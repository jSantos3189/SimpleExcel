namespace SimpleExcel.Attributes;
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnAttribute(string? name = null) : Attribute
{
    public string? Name { get; } = name;
    public int Order { get; set; } = int.MaxValue;
    public string? Format { get; set; }
}
