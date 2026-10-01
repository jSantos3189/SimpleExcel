namespace SimpleExcel.Options;
internal sealed class ExcelColumnOptions { public string? Header { get; set; } public bool Ignore { get; set; } public int? Order { get; set; } public string? Format { get; set; } }
public sealed class ExcelColumnBuilder
{
    private readonly ExcelColumnOptions _options; internal ExcelColumnBuilder(ExcelColumnOptions options)=>_options=options;
    public ExcelColumnBuilder Header(string text){ArgumentException.ThrowIfNullOrWhiteSpace(text);_options.Header=text;return this;}
    public ExcelColumnBuilder Ignore(bool value=true){_options.Ignore=value;return this;}
    public ExcelColumnBuilder Order(int order){if(order<0)throw new ArgumentOutOfRangeException(nameof(order));_options.Order=order;return this;}
    public ExcelColumnBuilder Format(string format){ArgumentException.ThrowIfNullOrWhiteSpace(format);_options.Format=format;return this;}
}
