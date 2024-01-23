using XLMapper.Core;
using XLMapper.Core.Models;

namespace XLMapper.Provider.FastExcel;

public class WorksheetProvider : IWorksheetProvider
{
    public Worksheet GetWorksheet(string filePath, int worksheetIndex = 1)
    {
        var inputFile = new FileInfo(filePath);
        using var fastExcel = new global::FastExcel.FastExcel(inputFile, true);

        var excelWorksheet = fastExcel.Read(worksheetIndex);
        var headers = GetHeaders(excelWorksheet);
        var rows = GetRows(headers, excelWorksheet);
        return new Worksheet
        {
            Name = excelWorksheet.Name,
            Headers = headers,
            Rows = rows.Skip(1).ToList()
        };
    }

    public Worksheet GetWorksheet(string filePath, string worksheetName)
    {
        var inputFile = new FileInfo(filePath);
        using var fastExcel = new global::FastExcel.FastExcel(inputFile, true);

        var excelWorksheet = fastExcel.Read(worksheetName);
        var headers = GetHeaders(excelWorksheet);
        var rows = GetRows(headers, excelWorksheet);
        return new Worksheet
        {
            Name = excelWorksheet.Name,
            Headers = headers,
            Rows = rows.Skip(1).ToList()
        };
    }

    private static IEnumerable<Row> GetRows(
        IReadOnlyCollection<Header> headers,
        global::FastExcel.Worksheet worksheet)
    {
        if (!worksheet.Rows.Any()) return [];

        return worksheet.Rows
            .Skip(1)
            .Select((r, i) =>
            {
                var cells = r.Cells
                    .ToDictionary(c => c.ColumnName, c => c.Value.ToString());

                return new Row
                {
                    RowNumber = i + 1,
                    Cells = headers
                        .Select(h => new Cell
                            {
                                ColumnName = h.Value!,
                                Value = cells.GetValueOrDefault(h.ColumnName)
                            }
                        )
                        .ToList()
                };
            }).ToList();
    }

    private static List<Header> GetHeaders(global::FastExcel.Worksheet worksheet)
    {
        if (!worksheet.Rows.Any()) return [];

        return worksheet.Rows
            .First().Cells
            .Select(c => new Header
            {
                ColumnName = c.ColumnName,
                Value = c.Value?.ToString() ?? string.Empty,
            })
            .Where(c => !string.IsNullOrWhiteSpace(c.Value))
            .ToList();
    }
}