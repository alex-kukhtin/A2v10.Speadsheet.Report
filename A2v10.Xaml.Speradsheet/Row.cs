using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace A2v10.Xaml.Spreadsheet;

[ContentProperty("Cells")]
public class Row : XamlElement
{
	public Object? ItemsSource { get; init; }
	public CellCollection Cells { get; init; } = new();

	internal void MergeStyles(StyleSet styleSet)
	{
		foreach (var cell in this.Cells)
			cell.MergeStyles(styleSet);
	}
}

public class RowCollection : List<Row>
{
}


