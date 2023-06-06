using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A2v10.Xaml.Speradsheet;

public record BindRuntime
{
	public String? Expression { get; init; }

	public String? Format { get; init; }
}
