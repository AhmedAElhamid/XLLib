using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using FastExcel;

namespace XLLib;

public class ExcelReader : IExcelReader
{
    public List<T> Read<T>(string filePath, int worksheetIndex = 1) where T : class, new() =>
        Read<T>(GetWorksheet(filePath, worksheetIndex));


    public List<T> Read<T>(string filePath, string worksheetName) where T : class, new()
        => Read<T>(GetWorksheet(filePath, worksheetName));


    private List<T> Read<T>(Worksheet worksheet) where T : class, new()
    {
        var rows = GetRows(worksheet);
        var properties = GetProperties<T>();
        var headers = GetHeadersAndRemoveItFromRows(rows);
        ValidateRequiredHeaders(headers, properties);
        return rows.Select((r, i) => GetItem<T>(r, headers, properties, i)).ToList();
    }


    private static void ValidateRequiredHeaders(string[] headers, PropertyInfo[] properties)
    {
        var requiredProperties = properties.Where(p => p.GetRequiredAttribute() != null).ToList();
        var requiredHeaders = headers.Where(h => requiredProperties.Any(p =>
            (p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name) == h)).ToList();
        var missingHeaders = requiredProperties.Where(p => !requiredHeaders.Contains(
            p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name)).ToList();
        if (missingHeaders.Count != 0)
            throw new ValidationException($"Missing headers: {string.Join(", ", missingHeaders.Select(p => p.Name))}");
    }

    private static string[] GetHeadersAndRemoveItFromRows(IList<string?[]> rows)
    {
        var headers = rows[0]
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Select(h => h!.Trim()).ToArray();
        rows.RemoveAt(0);
        return headers;
    }

    private static List<string?[]> GetRows(Worksheet worksheet)
    {
        if (!worksheet.Rows.Any()) return [];

        var headers = worksheet.Rows.First().Cells.Select(c => new
        {
            c.ColumnName,
            Value = c.Value.ToString(),
        }).ToArray();
        return worksheet.Rows
            .Select(r =>
            {
                if (r.Cells.Count() == headers.Length)
                    return r.Cells.Select(c => c.Value.ToString()).ToArray();

                var cells = r.Cells.Select(c => new
                {
                    c.ColumnName,
                    Value = c.Value.ToString(),
                }).ToArray();
                var cellsValues = headers
                    .Select(t => cells.Any(c => c.ColumnName == t.ColumnName)
                        ? cells.First(c => c.ColumnName == t.ColumnName).Value
                        : null)
                    .ToList();

                return cellsValues.ToArray();
            }).ToList();
    }

    private static PropertyInfo[] GetProperties<T>() where T : class, new()
        => typeof(T).GetProperties();

    private static T GetItem<T>(IReadOnlyList<string?> row,
        IReadOnlyList<string> headers, PropertyInfo[] properties, int rowIndex)
        where T : new()
    {
        var item = new T();
        for (var i = 0; i < headers.Count; i++)
        {
            var property = properties.FirstOrDefault(p =>
                (p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name) == headers[i]);
            if (property == null) continue;
            var value = row[i];
            SetProperty(property, item, value, rowIndex);
        }

        return item;
    }

    private static void SetProperty(PropertyInfo property, object item, string? value, int rowIndex)
    {
        if (property.PropertyType.IsGenericType &&
            property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            var underlyingType = Nullable.GetUnderlyingType(property.PropertyType);
            SetProperty(property, item, value, underlyingType!, rowIndex);
        }
        else
        {
            SetProperty(property, item, value, property.PropertyType, rowIndex);
        }
    }

    private static void SetProperty(PropertyInfo property, object item, string? value, Type type, int rowIndex)
    {
        if (type == typeof(string))
            property.SetValue(item, value.ThrowIfPropertyRequired(property, rowIndex));
        else if (type == typeof(int))
            property.SetValue(item, value.GetIntValue(rowIndex).ThrowIfPropertyRequired(property, rowIndex));
        else if (type == typeof(DateTime))
            property.SetValue(item, value.GetDateTimeValue(rowIndex).ThrowIfPropertyRequired(property, rowIndex));
        else if (type == typeof(decimal))
            property.SetValue(item, value.GetDecimalValue(rowIndex).ThrowIfPropertyRequired(property, rowIndex));
        else if (type == typeof(bool))
            property.SetValue(item, value.GetBoolValue(rowIndex).ThrowIfPropertyRequired(property, rowIndex));
        else if (type.IsEnum)
            property.SetValue(item, GetEnumValue(value, (Enum)property.GetValue(item)!)
                .ThrowIfPropertyRequired(property, rowIndex));
        else
        {
            throw new NotSupportedException(
                $"Property type {property.PropertyType} is not supported. Property name: {property.Name}");
        }
    }


    private static Worksheet GetWorksheet(string filePath, int worksheetIndex = 1)
    {
        var inputFile = new FileInfo(filePath);
        var fastExcel = new FastExcel.FastExcel(inputFile, true);

        return fastExcel.Read(worksheetIndex);
    }

    private static Worksheet GetWorksheet(string filePath, string worksheetName)
    {
        var inputFile = new FileInfo(filePath);
        using var fastExcel = new FastExcel.FastExcel(inputFile, true);

        return fastExcel.Read(worksheetName);
    }


    private static T? GetEnumValue<T>(string? value) where T : struct, Enum
    {
        return value == null ? null : (T)Enum.Parse(typeof(T), value);
    }

    private static T? GetNullableEnumValue<T>(string? value) where T : struct, Enum
    {
        return value == null ? null : (T?)Enum.Parse(typeof(T), value);
    }

    private static T? GetNullableEnumValue<T>(string? value, T? defaultValue) where T : struct, Enum
    {
        return value == null ? defaultValue : (T?)Enum.Parse(typeof(T), value);
    }

    private static T GetEnumValue<T>(string? value, T defaultValue) where T : Enum
    {
        return value == null ? defaultValue : (T)Enum.Parse(typeof(T), value);
    }
}