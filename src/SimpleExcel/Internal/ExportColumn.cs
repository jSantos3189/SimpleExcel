using System.Reflection; using SimpleExcel.Attributes; using SimpleExcel.Options;
namespace SimpleExcel.Internal;
internal sealed record ExportColumn(string Header,Type Type,Func<object,object?> Value,int Order,int NaturalOrder,string? Format);
internal static class ColumnFactory
{
 public static List<ExportColumn> For<T>(ExcelWriteOptions<T> options)=>typeof(T).GetProperties(BindingFlags.Public|BindingFlags.Instance)
 .Where(p=>p.CanRead&&p.GetCustomAttribute<ExcelIgnoreAttribute>() is null)
 .Select((p,n)=>{var a=p.GetCustomAttribute<ExcelColumnAttribute>();var c=options.Columns.GetValueOrDefault(p.Name);return new{P=p,A=a,C=c,N=n,O=c?.Order??a?.Order??int.MaxValue,F=c?.Format??a?.Format};})
 .Where(x=>x.C?.Ignore!=true).OrderBy(x=>x.O).ThenBy(x=>x.N)
 .Select(x=>new ExportColumn(x.C?.Header??x.A?.Name??x.P.Name,x.P.PropertyType,o=>x.P.GetValue(o),x.O,x.N,x.F)).ToList();
}
