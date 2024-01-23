using XLMapper.Core.Models;

namespace XLMapper.Core;

public interface IWorksheetProvider
{
    public Worksheet GetWorksheet(string filePath, int worksheetIndex = 1);
    public Worksheet GetWorksheet(string filePath, string worksheetName);
}