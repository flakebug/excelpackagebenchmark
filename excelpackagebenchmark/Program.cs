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
//			
//			//Array
//			xlsxbench = new benchmark();
//			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.EPPlus;
//			xlsxbench.DataSetMethod = benchmark.DataSetDefinition.Array;
//			xlsxbench.GenerateRandomArray(100000,20);
//			xlsxbench.WriteOperation();			
//			
//			//ConcurrentBag
//			xlsxbench = new benchmark();
//			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.EPPlus;
//			xlsxbench.DataSetMethod = benchmark.DataSetDefinition.ConcurrentBag;
//			xlsxbench.GenerateRandomConcurrentBag(100000,20);
//			xlsxbench.WriteOperation();
//			
//			//List
//			xlsxbench = new benchmark();
//			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.EPPlus;
//			xlsxbench.DataSetMethod = benchmark.DataSetDefinition.List;
//			xlsxbench.GenerateRandomList(100000,20);
//			xlsxbench.WriteOperation();
			
			//DataTable
			xlsxbench = new benchmark();
			xlsxbench.ExcelComponent = benchmark.ExcelComponentDefinition.EPPlus;
			xlsxbench.DataSetMethod = benchmark.DataSetDefinition.List;
			xlsxbench.GenerateRandomList(10000,20);
			xlsxbench.WriteOperation();		
			xlsxbench.ReadOperation();
			
			Console.WriteLine(">>> Excel components benchmark ends <<<");
			Console.WriteLine("Press any key to continue");
			Console.ReadKey(true);
		}
		

	}
}