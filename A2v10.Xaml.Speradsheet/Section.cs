using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace A2v10.Xaml.Spreadsheet;

[ContentProperty("Rows")]
public class Section : XamlElement
{
	public Object? ItemsSource { get; init; }
	public RowCollection Rows { get; init; } = new();
}
public class SectionCollection : List<Section>
{

}

