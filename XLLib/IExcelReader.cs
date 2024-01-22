namespace XLLib;

public interface IExcelReader
{
    public List<T> Read<T>(string filePath, int worksheetIndex = 1) where T : class, new();
    public List<T> Read<T>(string filePath, string worksheetName) where T : class, new();
}