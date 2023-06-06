using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace A2v10.Xaml.Speradsheet;

[ContentProperty("Rows")]
public class Sheet : XamlElement
{
	public String? Name { get; init; }
	public RowCollection Rows { get; init; } = new();
	public SectionCollection Sections { get; init; } = new();	
	public ColumnCollection Columns { get; init; } = new();	
	internal void MergeStyles(StyleSet styleSet)
	{
		foreach (var row in Rows)
			row.MergeStyles(styleSet);
		foreach (var section in Sections)	
			foreach (var row in section.Rows)
				row.MergeStyles(styleSet);
	}
}
public class SheetCollection : List<Sheet>
{

}

