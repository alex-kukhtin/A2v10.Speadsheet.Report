
using System;
using System.IO;
using System.Dynamic;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using A2v10.Data.Interfaces;
using A2v10.Data;
using A2v10.Xaml.Spreadsheet;
using System.Globalization;
using System.Linq;
using System.Xaml;

namespace A2v10.Spreadsheet.Report;

using RepWorkbook = A2v10.Xaml.Spreadsheet.Workbook;
using RepSheet = A2v10.Xaml.Spreadsheet.Sheet;
using RepRow = A2v10.Xaml.Spreadsheet.Row;
using RepCell = A2v10.Xaml.Spreadsheet.Cell;
using RepColumn = A2v10.Xaml.Spreadsheet.Column;
using DataType = A2v10.Xaml.Spreadsheet.DataType;

using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;
using Workbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using Sheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;

public class ExcelSpreadsheetGenerator
{

	public Stream FileToExcel(String path, IDataModel dataModel)
	{
		var wkbook = XamlServices.Load(path) as RepWorkbook
			?? throw new NullReferenceException(path);
		return WorkbookToExcel(wkbook, dataModel);
	}
	public Stream WorkbookToExcel(RepWorkbook book, IDataModel dataModel)
	{
		MemoryStream ms;
		ms = new MemoryStream();
		using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
		{
			WorkbookPart wbPart = doc.AddWorkbookPart();
			wbPart.Workbook = new Workbook();

			WorkbookStylesPart workStylePart = wbPart.AddNewPart<WorkbookStylesPart>();
			
			UInt32Value sheetId = 1;

			workStylePart.Stylesheet = AddStyles(book.CreateStyles());
			workStylePart.Stylesheet.Save();

			Sheets sheets = wbPart.Workbook.AppendChild<Sheets>(new Sheets());

			foreach (var sh in book.Sheets)
			{
				WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
				wsPart.Worksheet = CreateSheet(sh, dataModel); //new Worksheet(new SheetData());// CreateSheet(sh);

				Sheet sheet = new() { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++ , Name = sh.Name };
				sheets.Append(sheet);
			}

			wbPart.Workbook.Save();
			doc.Dispose();
		};
		ms.Seek(0, SeekOrigin.Begin);
		return ms;
	}

	Stylesheet AddStyles(StyleSet styles)
	{
		static Color autoColor() { return new Color() { Auto = true }; }

		Fonts fonts = new(
			new Font( // Index 0 - default
				new FontSize() { Val = 11 }

			),
			new Font( // Index 1 - bold
				new FontSize() { Val = 11 },
				new Bold()
			),
			new Font( // Index 2 - title
				new FontSize() { Val = 14 },
				new Bold()
			));

		Borders borders = new(
				new Border(), // index 0 default
				new Border( // index 1 black border
					new LeftBorder(autoColor()) { Style = BorderStyleValues.Thin },
					new RightBorder(autoColor()) { Style = BorderStyleValues.Thin },
					new TopBorder(autoColor()) { Style = BorderStyleValues.Thin },
					new BottomBorder(autoColor()) { Style = BorderStyleValues.Thin },
					new DiagonalBorder()),
				new Border( // index 2 bottom border
					new LeftBorder(),
					new RightBorder(),
					new TopBorder(),
					new BottomBorder(autoColor()) { Style = BorderStyleValues.Thin },
					new DiagonalBorder())
			);

		Fills fills = new(
				new Fill(new PatternFill() { PatternType = PatternValues.None }));

		NumberingFormats numFormats = new(
				/*date*/     new NumberingFormat() { FormatCode = "dd\\.mm\\.yyyy;@", NumberFormatId = 166 },
				/*datetime*/ new NumberingFormat() { FormatCode = "dd\\.mm\\.yyyy hh:mm;@", NumberFormatId = 167 },
				/*currency*/ new NumberingFormat() { FormatCode = "#,##0.00####;[Red]\\-#,##0.00####", NumberFormatId = 169 },
				/*number*/   new NumberingFormat() { FormatCode = "#,##0.######;[Red]-#,##0.######", NumberFormatId = 170 },
				/*time*/     new NumberingFormat() { FormatCode = "hh:mm;@", NumberFormatId = 165 }
			);

		CellFormats cellFormats = new(new CellFormat());

		for (var i = 1 /*1-based!*/; i < styles.Count; i++)
		{
			Style st = styles[i];
			cellFormats.Append(CreateCellFormat(st));
		}

		return new Stylesheet(numFormats, fonts, fills, borders, cellFormats);
	}


	CellFormat CreateCellFormat(Style st)
	{
		var cf = new CellFormat()
		{
			FontId = 0,
			ApplyAlignment = true,
			Alignment = new Alignment
			{
				Vertical = VerticalAlignmentValues.Top
			}
		};

		// dataType
		switch (st.DataType)
		{
			case DataType.Number:
				break;
			case DataType.Currency:
				cf.NumberFormatId = 169;
				cf.ApplyNumberFormat = true;
				break;
			case DataType.Date:
				cf.NumberFormatId = 166;
				cf.ApplyNumberFormat = true;
				cf.Alignment.Horizontal = HorizontalAlignmentValues.Center;
				break;
			case DataType.DateTime:
				cf.NumberFormatId = 167;
				cf.ApplyNumberFormat = true;
				cf.Alignment.Horizontal = HorizontalAlignmentValues.Center;
				break;
			case DataType.Time:
				cf.NumberFormatId = 165;
				cf.ApplyNumberFormat = true;
				break;
			case DataType.String:
				cf.Alignment.WrapText = true;
				cf.Alignment.Horizontal = HorizontalAlignmentValues.Left;
				cf.ApplyNumberFormat = false;
				break;
		}
		return cf;	
	}

	static Double ConvertUnit(UInt32 val)
	{
		const Decimal charWidth = 7;
		return (Double)Math.Truncate((val + 5L) / charWidth * 256L) / 256L;
	}

	void SetCellData(RepCell repCell, Cell cell, Object val)
	{
		if (val == null)
			return;
		switch (repCell.DataType)
		{
			case DataType.Date:
				if (val is DateTime dt)
				{
					// DataType not needed
					cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
				}
				break;
			case DataType.Number:
				cell.DataType = new EnumValue<CellValues>(CellValues.Number);
				cell.CellValue = new CellValue(Convert.ToDecimal(val));
				break;
			case DataType.Currency:
				cell.DataType = new EnumValue<CellValues>(CellValues.Number);
				cell.CellValue = new CellValue(Convert.ToDecimal(val));
				break;
			default:
				cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
				cell.InlineString = new InlineString(new Text(val.ToString()));
				break;
		}
	}

	Cell CreateCell(RepCell repCell, ExpandoObject source)
	{
		var cell = new Cell();
		if (repCell.StyleIndex > 0)
			cell.StyleIndex = repCell.StyleIndex;
		var content = repCell.GetBinding(nameof(repCell.Content));
		if (content != null)
		{
			var val = source.Eval<Object>(content.Expression);
			if (val != null)
			{
				SetCellData(repCell, cell, val);
			}

		}
		else if (repCell.Content != null)
		{
			SetCellData(repCell, cell, repCell.Content);
		}
		return cell;
	}

	Row CreateRow(RepRow repRow, ExpandoObject source)
	{
		var row = new Row();
		foreach (var shCell in repRow.Cells)
		{
			row.Append(CreateCell(shCell, source));
		}
		return row;
	}

	Worksheet CreateSheet(RepSheet sheet, IDataModel dataModel)
	{
		var sd = new SheetData();
		var cols = new Columns();

		var props = new SheetFormatProperties()
		{
			BaseColumnWidth = 10,
			DefaultRowHeight = 30,
			DyDescent = 0.25
		};

		foreach (var shSect in sheet.Sections)
		{
			var bindSect = shSect.GetBinding("ItemsSource");
			if (bindSect == null)
				continue;
			var source = dataModel.Eval<List<ExpandoObject>>(bindSect.Expression) 
				?? throw new InvalidOperationException($"Element not found '{bindSect.Expression}'");
			foreach (var elem in source)
			{
				foreach (var row in shSect.Rows)
				{
					var bindRow = row.GetBinding(nameof(row.ItemsSource));
					if (bindRow != null)
					{
						var rowSource = elem.Eval<List<ExpandoObject>>(bindRow.Expression);
						foreach (var chElem in rowSource)
							sd.Append(CreateRow(row, chElem));
					}
					else
						sd.Append(CreateRow(row, elem));
				}
			}
		}

		foreach (var shRow in sheet.Rows)
		{
			var bindRow = shRow.GetBinding(nameof(shRow.ItemsSource));
			if (bindRow != null)
			{
				// DataSource
				var source = dataModel.Eval<List<ExpandoObject>>(bindRow.Expression);
				foreach (var elem in source)
					sd.Append(CreateRow(shRow, elem));
			}
			else
				sd.Append(CreateRow(shRow, dataModel.Root));
		}

		for (UInt32 cx = 0; cx < sheet.Columns.Count; cx++)
		{
			var xcol = sheet.Columns[(int) cx];
			if (xcol.Width != 0)
				cols.Append(new Column() { Min = cx + 1, Max = cx + 1, BestFit = false, CustomWidth = true, Width = ConvertUnit(xcol.Width) });
		}
		//c++;
		//cols.Append(new Column() { Min = c + 1, Max = c + 1, BestFit = false, CustomWidth = true, Width = ConvertUnit(50) });
		/*
		if (_mergeCells.Count > 0)
		{
			var mc = new MergeCells();
			foreach (var mergeRef in _mergeCells)
				mc.Append(new MergeCell() { Reference = mergeRef });
			wsPart.Worksheet.Append(mc);
		}
		*/

		var ws = new Worksheet(props, sd);
		if (cols.Any())
		{
			ws.AddChild(cols);
		}

		ws.AddChild(new IgnoredErrors(
			new IgnoredError() { 
				NumberStoredAsText = true,
				SequenceOfReferences = new ListValue<StringValue>(
					new List<StringValue>() { new StringValue("A1:WZZ999999") })
			}
		));
		return ws;
	}
}
