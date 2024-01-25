// See https://aka.ms/new-console-template for more information


using XLMapper.Core;
using XLMapper.Provider.FastExcel;
using XLMapper.Test;

try
{
    var excelMapper = new ExcelMapper(new WorksheetProvider());
    var result = excelMapper.Map<Model>(Path.Combine("wwwroot", "input.xlsx"));
    Console.WriteLine(string.Join(Environment.NewLine, result));
}
catch (Exception e)
{
    Console.WriteLine(e);
    throw;
}