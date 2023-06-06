using System;
using System.Collections.Generic;

namespace A2v10.Xaml.Speradsheet;

public class Column : XamlElement
{
	public UInt32 Width { get; init; }
}

public class ColumnCollection : List<Column>
{

}


