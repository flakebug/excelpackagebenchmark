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
		private ExcelComponentDefinition _excelComponent;
		private int _dataRowCount;
		private int _dataColumnCount;
		
		Stopwatch _stopWatch;

		
		public enum ExcelComponentDefinition
		{
			None,
			EPPlus,
			NPOI,
			ClosedXML
			//SpreadsheetLight //expansion for future
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
//				//expansion for future					
//				case ExcelComponentDefinition.SpreadsheetLight:
//					spreadsheetlight_write();
//					break;
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
					closedxml_read();
					break;
				case ExcelComponentDefinition.NPOI:
					npoi_read();
					break;	
//				//Expansion for future					
//				case ExcelComponentDefinition.SpreadsheetLight:
//					spreadsheetlight_write();
//					break;
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
		private void npoi_read()
		{
			const string filename = "npoi.xlsx";
			FileInfo file = new FileInfo(filename);
			
			_stopWatch.Reset();
			_stopWatch.Start();				
			var workbook = new XSSFWorkbook(file);
			_stopWatch.Stop();
			PrintTime("npoi_read() - read filestream to npoi object");	
			
			_stopWatch.Reset();
			_stopWatch.Start();					
			var sht = workbook.GetSheet(worksheetname);
			_stopWatch.Stop();
			PrintTime("npoi_read() - assign npoi.worksheet to variable");
			
			//test the performance of filling cells				
			_stopWatch.Reset();
			_stopWatch.Start();
			Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
				//DataRow dr = _dataTable.Rows[rowIndex];
				for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
					lock (sht) {
						_dataTable.Rows[rowIndex][colIndex] = sht.GetRow(rowIndex).GetCell(colIndex).ToString();
					}
				}
			});			
			_stopWatch.Stop();
			PrintTime("npoi_read() - read excel to memory");
			
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
		private void closedxml_read()
		{
			const string filename = "closedxml.xlsx";

			_stopWatch.Reset();
			_stopWatch.Start();	
			var workbook = new XLWorkbook(filename);
			_stopWatch.Stop();
			PrintTime("closedxml_read() - read filestream to EPPlus object");

			_stopWatch.Reset();
			_stopWatch.Start();				
			var sht = workbook.Worksheets.Worksheet(worksheetname);
			_stopWatch.Stop();
			PrintTime("closedxml_read() - assign ClosedXML.worksheet to variable");
			
			//test the performance of filling cells				
			_stopWatch.Reset();
			_stopWatch.Start();					
			Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
				//DataRow dr = _dataTable.Rows[rowIndex];
				for (int colIndex = 0;
					     colIndex < _dataTable.Columns.Count;
					     colIndex++) {
					lock (_dataTable.Rows) {
						_dataTable.Rows[rowIndex][colIndex] = sht.Cell(rowIndex + 1, colIndex + 1).Value.ToString();
					}
				}
			});			
			_stopWatch.Stop();
			PrintTime("closedxml_read() - read to memory");
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
				Parallel.For(0, _dataTable.Rows.Count, rowIndex => {
					DataRow dr = _dataTable.Rows[rowIndex];
					for (int colIndex = 0;
							     colIndex < _dataTable.Columns.Count;
							     colIndex++) {
						sht.Cells[rowIndex + 1, colIndex + 1].Value = dr[colIndex];
					}
				});	

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
