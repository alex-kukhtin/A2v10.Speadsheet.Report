using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A2v10.Xaml.Spreadsheet;

public class XamlElement : ISupportBinding
{
	BindImpl? _bindImpl;
	#region ISupportBinding
	public BindImpl BindImpl
	{
		get
		{
			_bindImpl ??= new BindImpl();
			return _bindImpl;
		}
	}
	public BindRuntime? GetBinding(String name)
	{
		return _bindImpl?.GetBinding(name)?.Runtime();
	}
	#endregion
}
