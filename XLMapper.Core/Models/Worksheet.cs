namespace XLMapper.Core.Models;

public class Worksheet
{
    public string Name { get; set; } = string.Empty;
    public List<Header> Headers { get; set; } = new();
    public List<Row> Rows { get; set; } = new();
}