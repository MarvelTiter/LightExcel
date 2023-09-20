using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Utils
{
    internal class ReflectGetter
    {
        private readonly Func<object, object> func;

        public ReflectGetter(PropertyInfo property)
        {
            func = CreateGetterDelegate(property);
        }

        public object Invoke(object instance)
        {
            return func.Invoke(instance);
        }

        private static Func<object, object> CreateGetterDelegate(PropertyInfo property)
        {
            var pExp = Expression.Parameter(typeof(object));
            var typedPExp = Expression.Convert(pExp, property.DeclaringType!);
            var propExp = Expression.Property(typedPExp, property);
            var ret = Expression.Convert(propExp, typeof(object));
            return Expression.Lambda<Func<object, object>>(ret, pExp).Compile();
        }
    }
    public class Property
    {
        public PropertyInfo Info { get; }
        ReflectGetter? getter;
        public Property(PropertyInfo info)
        {
            Info = info;
        }
        public bool CanRead => Info.CanRead;
        public object GetValue(object instance)
        {
            if (CanRead)
            {
                getter ??= new ReflectGetter(Info);
                return getter.Invoke(instance);
            }
            throw new NotSupportedException();
        }
    }
}
