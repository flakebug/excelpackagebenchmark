/*
 * Created by SharpDevelop.
 * User: 53785
 * Date: 2017/12/27
 * Time: 上午 08:27
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Diagnostics;

namespace excelpackagebenchmark
{
	class Program
	{
		public static void Main(string[] args)
		{
			Console.WriteLine(">>> Excel components benchmark starts <<<");
			benchmark xlsxbench;
			
			const int dataRowCount = 100000;
			const int dataColCount = 20;

			xlsxbench = new benchmark();
			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.NPOI;
			xlsxbench.GenerateRandomDataTable(dataRowCount, dataColCount);
			xlsxbench.WriteOperation();		
			xlsxbench.ReadOperation();

			xlsxbench = new benchmark();
			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.EPPlus;
			xlsxbench.GenerateRandomDataTable(dataRowCount, dataColCount);
			xlsxbench.WriteOperation();		
			xlsxbench.ReadOperation();
			
			xlsxbench = new benchmark();
			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.ClosedXML;
			xlsxbench.GenerateRandomDataTable(dataRowCount, dataColCount);
			xlsxbench.WriteOperation();		
			xlsxbench.ReadOperation();
			
			Console.WriteLine(">>> Excel components benchmark ends <<<");
			Console.WriteLine("Press any key to continue");
			Console.ReadKey(true);
		}
		

	}
}