using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A2v10.Xaml.Speradsheet;

public interface ISupportBinding
{
	BindImpl BindImpl { get; }
	BindRuntime? GetBinding(String name);
}
