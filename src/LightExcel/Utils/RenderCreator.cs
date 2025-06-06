using System.Linq.Expressions;

namespace LightExcel.Utils;

internal static class RenderCreator<T> where T : IDataRender
{
    static readonly Func<ExcelConfiguration, IDataRender> function;
    static RenderCreator()
    {
        var ctor = typeof(T).GetConstructors().First(m => m.GetParameters().Length == 1);
        ParameterExpression p = Expression.Parameter(typeof(ExcelConfiguration), "config");
        var body = Expression.New(ctor, p);
        var lambda = Expression.Lambda<Func<ExcelConfiguration, IDataRender>>(body, p);
        function = lambda.Compile();
    }

    public static IDataRender Create(ExcelConfiguration config) => function.Invoke(config);
}
#if NET6_0_OR_GREATER
internal static class AsyncRenderCreator<TRender> where TRender : IAsyncDataRender
{
    static readonly Func<ExcelConfiguration, TRender> function;
    static AsyncRenderCreator()
    {
        var ctor = typeof(TRender).GetConstructors().First(m => m.GetParameters().Length == 1);
        ParameterExpression p = Expression.Parameter(typeof(ExcelConfiguration), "config");
        var body = Expression.New(ctor, p);
        var lambda = Expression.Lambda<Func<ExcelConfiguration, TRender>>(body, p);
        function = lambda.Compile();
    }

    public static TRender Create(ExcelConfiguration config) => function.Invoke(config);
}
#endif
