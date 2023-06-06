using System;
using System.Collections.Generic;

namespace A2v10.Xaml.Speradsheet;

public class Cell : ContentElement
{
	public UInt32 StyleIndex = 0;
	public DataType DataType { get; init; }

	internal void MergeStyles(StyleSet styleSet)
	{
		var bind = GetBinding(nameof(Content));
		if (bind != null)
		{
			StyleIndex = styleSet.AddStyle(new Style(DataType));
		}
	}
}

public class CellCollection : List<Cell>
{

}


