namespace SimpleExcel.Models;
public sealed record ExcelReadError(uint Row, string Column, string Message);
public sealed class ExcelReadResult<T>
{
    public List<T> Data { get; } = [];
    public List<ExcelReadError> Errors { get; } = [];
    public bool IsValid => Errors.Count == 0;
}
