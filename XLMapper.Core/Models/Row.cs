namespace XLMapper.Core.Models;

public class Row
{
    public int RowNumber { get; set; }
    public List<Cell> Cells { get; set; } = new();
}