/*
 * Created by SharpDevelop.
 * User: 53785
 * Date: 2017/12/27
 * Time: 上午 08:45
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Diagnostics;
using OfficeOpenXml;
using ClosedXML.Excel;
using NPOI.XSSF.UserModel;


namespace excelpackagebenchmark
{
	/// <summary>
	/// Description of benchmark.
	/// </summary>
	public class benchmark
	{
		private const string worksheetname = "worksheet";
		private DataTable _dataTable;
		private ConcurrentBag<CellDefinition> _concurrentBagDataSet;
		private List<CellDefinition> _listDataSet;
		private CellDefinition[] _arrayDataSet;
		private DataSetDefinition _datasetMethod;
		private ExcelComponentDefinition _excelComponent;
		private int _dataRowCount;
		private int _dataColumnCount;
		
		Stopwatch _stopWatch;
		
		public enum DataSetDefinition
		{
			DataTable,
			ConcurrentBag,
			List,
			Array
		}
		
		public enum ExcelComponentDefinition
		{
			None,
			EPPlus,
			NPOI,
			SpreadsheetLight,
			ClosedXML
		}
		
		public struct CellDefinition
		{
			public int RowNumber;
			public int ColumnNumber;
			public string Text;
		}
		
		public ExcelComponentDefinition ExcelComponent {
			get { return _excelComponent; }
			set { _excelComponent = value; }
		}
		public DataSetDefinition DataSetMethod {
			get { return _datasetMethod; }
			set { _datasetMethod = value; }
		}
		public DataTable DataTable {
			get { return _dataTable; }
			set { _dataTable = value; }
		}
		
		public benchmark()
		{
			_stopWatch = new Stopwatch();
		}
		
		public void WriteOperation()
		{
			switch (_excelComponent) {
				case ExcelComponentDefinition.None:
					throw new Exception("Excel component is not assigned");
					break;
				case ExcelComponentDefinition.EPPlus:
					epplus_write();
					break;
				case ExcelComponentDefinition.ClosedXML:
					closedxml_write();
					break;
				case ExcelComponentDefinition.NPOI:
					npoi_write();
					break;					
				case ExcelComponentDefinition.SpreadsheetLight:
					//spreadsheetlight_write();
					break;
			}
		}
		public void ReadOperation()
		{
			switch (_excelComponent) {
				case ExcelComponentDefinition.None:
					throw new Exception("Excel component is not assigned");
					break;
				case ExcelComponentDefinition.EPPlus:
					epplus_read();
					break;
				case ExcelComponentDefinition.ClosedXML:
					//closedxml_write();
					break;
				case ExcelComponentDefinition.NPOI:
					//npoi_write();
					break;					
				case ExcelComponentDefinition.SpreadsheetLight:
					//spreadsheetlight_write();
					break;
			}
		}

		private void npoi_write()
		{
			const string filename = "npoi.xlsx";
			if (File.Exists(filename)) {
				File.Delete(filename);
			}			
			var workbook = new XSSFWorkbook();
			var sht = workbook.CreateSheet(worksheetname);			
			
			//test the performance of filling cells				
			_stopWatch.Reset();
			_stopWatch.Start();
			for (int rowIndex = 0;
			     rowIndex < _dataTable.Rows.Count;
			     rowIndex++) {
				sht.CreateRow(rowIndex);
			}
			Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
				DataRow dr = _dataTable.Rows[rowIndex];
				for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
					lock (sht) {
						sht.GetRow(rowIndex).CreateCell(colIndex).SetCellValue(dr[colIndex].ToString());
					}
				}
			});			
			_stopWatch.Stop();
			PrintTime("npoi_write() - Fill Cells");
			
			//test the performance of write file
			_stopWatch.Reset();
			_stopWatch.Start();				
			FileStream file = new FileStream(filename, FileMode.Create);
			workbook.Write(file);
			file.Close();
			PrintTime("npoi_write() - Save file to disk");			
		}
		private void closedxml_write()
		{
			const string filename = "closedxml.xlsx";
			if (File.Exists(filename)) {
				File.Delete(filename);
			}			
			var workbook = new XLWorkbook();
			var sht = workbook.Worksheets.Add(worksheetname);
			//test the performance of filling cells				
			_stopWatch.Reset();
			_stopWatch.Start();					
			Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
				DataRow dr = _dataTable.Rows[rowIndex];
				for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
					lock (sht) {
						sht.Cell(rowIndex + 1, colIndex + 1).Value = dr[colIndex];
					}
				}
			});			
			_stopWatch.Stop();
			PrintTime("closedxml_write() - Fill Cells");
			
			//test the performance of write file
			_stopWatch.Reset();
			_stopWatch.Start();				
			workbook.SaveAs(filename);
			PrintTime("closedxml_write() - Save file to disk");			
		}
		private void epplus_write()
		{
			const string filename = "epplus.xlsx";
			if (File.Exists(filename)) {
				File.Delete(filename);
			}
			FileInfo xlsx = new FileInfo(filename);
			using (ExcelPackage epplus = new ExcelPackage(xlsx)) {   
				ExcelWorksheet sht = epplus.Workbook.Worksheets.Add(worksheetname);		

				//test the performance of filling cells				
				_stopWatch.Reset();
				_stopWatch.Start();			
				switch (DataSetMethod) {
					case DataSetDefinition.Array:
						Parallel.For(0, _arrayDataSet.Length, (index) => {
							sht.Cells[_arrayDataSet[index].RowNumber, _arrayDataSet[index].ColumnNumber].Value = _arrayDataSet[index].Text;
						});								
						break;
					case DataSetDefinition.List:
						Parallel.ForEach(_listDataSet, (item) => {
							sht.Cells[item.RowNumber, item.ColumnNumber].Value = item.Text;
						});						
						break;
					case DataSetDefinition.ConcurrentBag:
						Parallel.ForEach(_concurrentBagDataSet, (item) => {
							sht.Cells[item.RowNumber, item.ColumnNumber].Value = item.Text;
						});
						break;
					case DataSetDefinition.DataTable:
						Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
							DataRow dr = _dataTable.Rows[rowIndex];
							for (int colIndex = 0;
							     colIndex < _dataTable.Columns.Count;
							     colIndex++) {
								sht.Cells[rowIndex + 1, colIndex + 1].Value = dr[colIndex];
							}
						});	
						break;
				}

				_stopWatch.Stop();
				PrintTime("epplus_write() - Fill Cells");
				
				//test the performance of write file
				_stopWatch.Reset();
				_stopWatch.Start();					
				epplus.Save();
				_stopWatch.Stop();
				PrintTime("epplus_write() - Save file to disk");				
			}
		}
		private void epplus_read()
		{
			const string filename = "epplus.xlsx";
			if (!File.Exists(filename)) {
				throw new Exception("file not found");
			}
			
			FileInfo xlsx = new FileInfo(filename);
			
			_stopWatch.Reset();
			_stopWatch.Start();			
			ExcelPackage epplus = new ExcelPackage(xlsx);
			_stopWatch.Stop();
			PrintTime("epplus_read() - read filestream to EPPlus object");
			
			_stopWatch.Reset();
			_stopWatch.Start();	
			ExcelWorksheet sht = epplus.Workbook.Worksheets[worksheetname];
			_stopWatch.Stop();
			PrintTime("epplus_read() - assign EPPlus.worksheet to variable");	
			
			_stopWatch.Reset();
			_stopWatch.Start();				
			switch (DataSetMethod) {
				case DataSetDefinition.DataTable:
					Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
						DataRow dr = _dataTable.Rows[rowIndex];
						for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
							lock (_dataTable) {
								_dataTable.Rows[rowIndex][colIndex] = sht.Cells[rowIndex + 1, colIndex + 1].Text;
							}
						}
					});	
					break;
				case DataSetDefinition.ConcurrentBag:
					_concurrentBagDataSet = new ConcurrentBag<CellDefinition>();
					//test the performance of reading cells				
					Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
						for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
							CellDefinition cl;
							cl.RowNumber = rowIndex;
							cl.ColumnNumber = colIndex;
							cl.Text = sht.Cells[rowIndex + 1, colIndex + 1].Text;
							_concurrentBagDataSet.Add(cl);
						}
					});	
					break;
				case DataSetDefinition.List:
					_listDataSet = new List<CellDefinition>();
					//test the performance of reading cells				
					Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
						for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
							CellDefinition cl;
							cl.RowNumber = rowIndex;
							cl.ColumnNumber = colIndex;
							cl.Text = sht.Cells[rowIndex + 1, colIndex + 1].Text;
							lock(_listDataSet) {
								_listDataSet.Add(cl);
							}
						}
					});	
					break;					
			}
			_stopWatch.Stop();
			PrintTime("epplus_read() - read file to memory");

			epplus.Dispose();
		}
		public void GenerateRandomDataTable(int RowCount, int ColumnCount)
		{
			_stopWatch.Reset();
			_stopWatch.Start();					
			_dataTable = new DataTable();
			for (int colIndx = 0;
			     colIndx < ColumnCount;
			     colIndx++) {
				_dataTable.Columns.Add("c" + colIndx, typeof(string));
				
			}

			Parallel.For(0, RowCount, rowIndex => {
				DataRow dr;
				lock (_dataTable) {
					dr = _dataTable.NewRow();
				}
				for (int colIndex = 0;
			     colIndex < ColumnCount;
			     colIndex++) {
					dr[colIndex] = Guid.NewGuid().ToString();
				}
				lock (_dataTable) {
					_dataTable.Rows.Add(dr);
				}
			});
			_stopWatch.Stop();
			PrintTime("GenerateRandomDataTable()");	
		}
		public void GenerateRandomConcurrentBag(int RowCount, int ColumnCount)
		{
			_stopWatch.Reset();
			_stopWatch.Start();					
			_concurrentBagDataSet = new ConcurrentBag<CellDefinition>();

			Parallel.For(0, RowCount, rowIndex => {
				for (int colIndex = 0;
			     colIndex < ColumnCount;
			     colIndex++) {
					CellDefinition cl = new CellDefinition();
					cl.RowNumber = rowIndex + 1;
					cl.ColumnNumber = colIndex + 1;
					cl.Text = Guid.NewGuid().ToString();
					_concurrentBagDataSet.Add(cl);
				}
			});
			_stopWatch.Stop();
			PrintTime("GenerateRandomConcurrentBag()");						
					
		}
		public void GenerateRandomList(int RowCount, int ColumnCount)
		{
			_stopWatch.Reset();
			_stopWatch.Start();					
			_listDataSet = new List<CellDefinition>();

			Parallel.For(0, RowCount, rowIndex => {
				for (int colIndex = 0;
			     colIndex < ColumnCount;
			     colIndex++) {
					CellDefinition cl = new CellDefinition();
					cl.RowNumber = rowIndex + 1;
					cl.ColumnNumber = colIndex + 1;
					cl.Text = Guid.NewGuid().ToString();
					lock (_listDataSet) {
						_listDataSet.Add(cl);
					}
				}
			});
			_stopWatch.Stop();
			PrintTime("GenerateRandomList()");						
					
		}
		public void GenerateRandomArray(int RowCount, int ColumnCount)
		{
			_stopWatch.Reset();
			_stopWatch.Start();					
			List<CellDefinition> _localListDataSet = new List<CellDefinition>();

			Parallel.For(0, RowCount, rowIndex => {
				for (int colIndex = 0;
			     colIndex < ColumnCount;
			     colIndex++) {
					CellDefinition cl = new CellDefinition();
					cl.RowNumber = rowIndex + 1;
					cl.ColumnNumber = colIndex + 1;
					cl.Text = Guid.NewGuid().ToString();
					lock (_localListDataSet) {
						_localListDataSet.Add(cl);
					}
				}
			});
			_arrayDataSet = _localListDataSet.ToArray();
			_stopWatch.Stop();
			PrintTime("GenerateRandomArray()");						
					
		}
		public void PrintTime(string message)
		{
			TimeSpan ts = _stopWatch.Elapsed;
			string elapsedTime = String.Format("{0} - {1:00}.{2:00} seconds",
				                     message,
				                     ts.Seconds,
				                     ts.Milliseconds / 10);				
			Console.WriteLine(elapsedTime); 
		}
	}
}
