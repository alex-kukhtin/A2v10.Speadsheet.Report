// See https://aka.ms/new-console-template for more information

using A2v10.Data;
using A2v10.Speadsheet.Report;
using A2v10.Xaml.Speradsheet;
using System;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Xaml;

namespace TestApp;
internal static class Program
{

	static void DeleteFile(String path)
	{
		try
		{
			File.Delete(path);
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex.Message);
		}
	}

	static void Main()
	{
		Console.WriteLine("Start");

		var path = "C:\\A2v10_Net48\\A2v10.Speadsheet.Report\\TestApp\\excelReport.xaml";
		var wkbook = XamlServices.Load(path) as Workbook
			?? throw new NullReferenceException(path);	

		Console.WriteLine(wkbook);

		var loc = new NullDataLocalizer();
		var prof = new NullDataProfiler();
		var config = new SimpleDataConfiguration();

		var dbContext = new SqlDbContext(prof, config, loc);

		var dm = dbContext.LoadModel(null, "ee.[Order.Index.Export]", new ExpandoObject()
		{
			{ "UserId", 99 },
			//{ "Id", 354 }
			{ "Month", "20230401" }
		});
		Console.WriteLine(dm);

		var gen = new ExcelGenerator();
		var ms = gen.WorkbookToExcel(wkbook, dm);

		var outPath = "d:\\temp\\sample.xlsx";
		DeleteFile(outPath);
		using (var outFile = File.OpenWrite(outPath))
		{
			ms.CopyTo(outFile);
		}

		Process.Start("excel.exe", outPath);
	}
}
