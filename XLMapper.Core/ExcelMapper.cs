using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using XLMapper.Core.Models;

namespace XLMapper.Core;

public class ExcelMapper : IExcelMapper
{
    private readonly IWorksheetProvider _worksheetProvider;

    public ExcelMapper(IWorksheetProvider worksheetProvider)
    {
        _worksheetProvider = worksheetProvider;
    }

    public List<T> Map<T>(string filePath, int worksheetIndex = 1) where T : class, new() =>
        Map<T>(_worksheetProvider.GetWorksheet(filePath, worksheetIndex));


    public List<T> Map<T>(string filePath, string worksheetName) where T : class, new()
        => Map<T>(_worksheetProvider.GetWorksheet(filePath, worksheetName));


    private static List<T> Map<T>(Worksheet worksheet) where T : class, new()
    {
        var properties = GetProperties<T>();
        ValidateRequiredHeaders(worksheet.Headers, properties);
        return worksheet.Rows
            .Select(r => GetItem<T>(r, worksheet.Headers, properties)).ToList();
    }


    private static void ValidateRequiredHeaders(List<Header> headers, IEnumerable<PropertyInfo> properties)
    {
        var requiredProperties = properties.Where(p => p.GetRequiredAttribute() != null).ToList();
        var requiredHeaders = headers.Where(h => requiredProperties.Any(p =>
            (p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name) == h.Value)).ToList();
        var missingHeaders = requiredProperties.Where(p =>
                requiredHeaders.All(h =>
                    h.Value != (p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name)))
            .ToList();
        if (missingHeaders.Count != 0)
            throw new ValidationException($"Missing header/s: {string.Join(", ", missingHeaders.Select(p => p.Name))}");
    }

    private static PropertyInfo[] GetProperties<T>() where T : class, new()
        => typeof(T).GetProperties();

    private static T GetItem<T>(Row row,
        IReadOnlyList<Header> headers, PropertyInfo[] properties)
        where T : new()
    {
        var item = new T();
        for (var i = 0; i < headers.Count; i++)
        {
            var property = properties.FirstOrDefault(p =>
                (p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name) == headers[i].Value);
            if (property == null) continue;
            var value = row.Cells[i].Value;
            SetProperty(property, item, value, row.RowNumber);
        }

        return item;
    }

    private static void SetProperty(PropertyInfo property, object item, string? value, int rowNumber)
    {
        if (property.PropertyType.IsGenericType &&
            property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            var underlyingType = Nullable.GetUnderlyingType(property.PropertyType);
            SetProperty(property, item, value, underlyingType!, rowNumber);
        }
        else
        {
            SetProperty(property, item, value, property.PropertyType, rowNumber);
        }
    }

    private static void SetProperty(PropertyInfo property, object item, string? value, Type type, int rowNumber)
    {
        if (type == typeof(string))
            property.SetValue(item, value.ThrowIfPropertyRequired(property, rowNumber));
        else if (type == typeof(int))
            property.SetValue(item, value.GetIntValue(rowNumber).ThrowIfPropertyRequired(property, rowNumber));
        else if (type == typeof(DateTime))
            property.SetValue(item, value.GetDateTimeValue(rowNumber).ThrowIfPropertyRequired(property, rowNumber));
        else if (type == typeof(decimal))
            property.SetValue(item, value.GetDecimalValue(rowNumber).ThrowIfPropertyRequired(property, rowNumber));
        else if (type == typeof(bool))
            property.SetValue(item, value.GetBoolValue(rowNumber).ThrowIfPropertyRequired(property, rowNumber));
        else if (type.IsEnum)
            property.SetValue(item, value.GetEnumValue((Enum)property.GetValue(item)!)
                .ThrowIfPropertyRequired(property, rowNumber));
        else
        {
            throw new NotSupportedException(
                $"Property type {property.PropertyType} is not supported. Property name: {property.Name}");
        }
    }
}