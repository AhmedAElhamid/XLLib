namespace XLMapper.Core;

public interface IExcelMapper
{
    public List<T> Map<T>(string filePath, int worksheetIndex = 1) where T : class, new();
    public List<T> Map<T>(string filePath, string worksheetName) where T : class, new();
}