using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace A2v10.Xaml.Spreadsheet;

[ContentProperty("Content")]
public class ContentElement : XamlElement
{
	public Object? Content { get; init; }
}
