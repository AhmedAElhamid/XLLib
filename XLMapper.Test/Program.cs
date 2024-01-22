// See https://aka.ms/new-console-template for more information


using XLMapper.Core;
using XLMapper.Test;

try
{
    IExcelReader excelReader = new ExcelReader();
    var result = excelReader.Read<Model>("input.xlsx");
    Console.WriteLine("Hello, World!");
}
catch (Exception e)
{
    Console.WriteLine(e);
    throw;
}

