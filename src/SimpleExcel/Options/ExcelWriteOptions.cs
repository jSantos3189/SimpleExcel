using System.Linq.Expressions;
namespace SimpleExcel.Options;
public class ExcelWriteOptions
{
 public string SheetName { get; set; }="Datos"; public bool IncludeHeaders { get; set; }=true; public bool AutoFitColumns { get; set; }=true; public bool WrapText { get; set; } public bool FreezeHeader { get; set; }=true;
 public ExcelHeaderOptions Header { get; }=new(); public ExcelColumnWidthOptions ColumnWidth { get; }=new();
}
public sealed class ExcelWriteOptions<T>:ExcelWriteOptions
{
 internal Dictionary<string,ExcelColumnOptions> Columns { get; }=new(StringComparer.Ordinal);
 public ExcelColumnBuilder Column<TProperty>(Expression<Func<T,TProperty>> expression){var body=expression.Body is UnaryExpression u?u.Operand:expression.Body;if(body is not MemberExpression m)throw new ArgumentException("La expresión debe seleccionar una propiedad.",nameof(expression));if(!Columns.TryGetValue(m.Member.Name,out var o))Columns[m.Member.Name]=o=new ExcelColumnOptions();return new ExcelColumnBuilder(o);}
}
