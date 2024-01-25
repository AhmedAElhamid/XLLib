// See https://aka.ms/new-console-template for more information


using XLMapper.Core;
using XLMapper.Provider;
using XLMapper.Provider.FastExcel;
using XLMapper.Test;

try
{
    IExcelMapper excelMapper = new ExcelMapper(new WorksheetProvider());
    var result = excelMapper.Map<Model>("input.xlsx");
    Console.WriteLine("Hello, World!");
}
catch (Exception e)
{
    Console.WriteLine(e);
    throw;
}

