using System;
using System.Windows.Markup;

namespace A2v10.Xaml.Spreadsheet;

[ContentProperty("Sheets")]
public class Workbook : XamlElement
{
	public SheetCollection Sheets { get; init; } = new();

	public StyleSet CreateStyles()
	{
		var styleSet = new StyleSet
		{
			new Style()
		};
		foreach (var sheet in Sheets) 
		{
			sheet.MergeStyles(styleSet);
		}
		return styleSet;	
	}
}
