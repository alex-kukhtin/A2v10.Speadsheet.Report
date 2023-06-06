using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A2v10.Xaml.Speradsheet;

public record struct Style
{
	public Style(DataType dataType)
	{
		DataType = dataType;	
	}
	public DataType DataType { get; set; }
}

public class StyleSet : List<Style>
{
	public UInt32 AddStyle(Style style)
	{
		var found = FindIndex(s => s.Equals(style));
		if (found != -1)
			return (UInt32) found;
		Add(style);
		return (UInt32)(Count - 1);
	}
}

