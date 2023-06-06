using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace A2v10.Xaml.Speradsheet;

public class Bind : MarkupExtension
{
	private BindImpl? _bindImpl;

	public BindImpl BindImpl
	{
		get
		{
			_bindImpl ??= new BindImpl();
			return _bindImpl;
		}
	}

	public Bind()
	{
	}

	public Bind(String path)
	{
		Expression = path;
	}

	public String? Expression { get; init; }

	public String? Format { get; init; }

	public BindRuntime Runtime()
	{
		return new BindRuntime()
		{
			Expression = this.Expression,
			Format = this.Format
		};
	}

	public override Object? ProvideValue(IServiceProvider serviceProvider)
	{
		if (serviceProvider.GetService(typeof(IProvideValueTarget)) is not IProvideValueTarget iTarget)
			return null;
		if (iTarget.TargetProperty is not PropertyInfo targetProp)
			return null;
		if (iTarget.TargetObject is not ISupportBinding targetObj)
			return null;
		targetObj.BindImpl.SetBinding(targetProp.Name, this);
		if (targetProp.PropertyType.IsValueType)
			return Activator.CreateInstance(targetProp.PropertyType);
		return null; // is object
	}

	public Bind? GetBinding(String name)
	{
		return _bindImpl?.GetBinding(name);
	}

	void SetBinding(String name, Bind bind)
	{
		BindImpl.SetBinding(name, bind);
	}
}
