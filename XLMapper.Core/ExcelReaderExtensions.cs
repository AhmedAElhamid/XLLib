using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace XLMapper.Core;

public static class ExcelReaderExtensions
{
    public static int? GetIntValue(this string? value, int rowIndex)
    {
        if (string.IsNullOrEmpty(value)) return null;
        return int.TryParse(value, out var intValue)
            ? intValue
            : throw new ValidationException($"Value {value} is not a valid integer at row: {rowIndex + 1}");
    }

    public static DateTime? GetDateTimeValue(this string? value, int rowIndex)
    {
        if (string.IsNullOrEmpty(value)) return null;
        if (double.TryParse(value, out var doubleValue)) return DateTime.FromOADate(doubleValue);
        return DateTime.TryParse(value, out var dateTimeValue)
            ? dateTimeValue
            : throw new ValidationException($"Value {value} is not a valid date at row: {rowIndex + 1}");
    }

    public static decimal? GetDecimalValue(this string? value, int rowIndex)
    {
        if (string.IsNullOrEmpty(value)) return null;
        return decimal.TryParse(value, out var decimalValue)
            ? decimalValue
            : throw new ValidationException($"Value {value} is not a valid decimal at row: {rowIndex + 1}");
    }

    public static bool? GetBoolValue(this string? value, int rowIndex)
    {
        if (string.IsNullOrEmpty(value)) return null;
        return value switch
        {
            "1" => true,
            "0" => false,
            _ => bool.TryParse(value, out var boolValue)
                ? boolValue
                : throw new ValidationException($"Value {value} is not a valid boolean at row: {rowIndex + 1}")
        };
    }

    public static string? ThrowIfPropertyRequired(
        this string? value,
        MemberInfo property,
        int rowIndex)
    {
        if (property.GetRequiredAttribute() == null || !string.IsNullOrWhiteSpace(value)) return value;
        throw new ValidationException($"Property {property.Name} is required. At row: {rowIndex + 1}");
    }

    public static int? ThrowIfPropertyRequired(
        this int? value,
        MemberInfo property,
        int rowIndex)
    {
        if (property.GetRequiredAttribute() == null || value != null) return value;
        throw new ValidationException($"Property {property.Name} is required. At row: {rowIndex + 1}");
    }

    public static DateTime? ThrowIfPropertyRequired(
        this DateTime? value,
        MemberInfo property,
        int rowIndex)
    {
        if (property.GetRequiredAttribute() == null || value != null) return value;
        throw new ValidationException($"Property {property.Name} is required. At row: {rowIndex + 1}");
    }

    public static decimal? ThrowIfPropertyRequired(
        this decimal? value,
        MemberInfo property,
        int rowIndex)
    {
        if (property.GetRequiredAttribute() == null || value != null) return value;
        throw new ValidationException($"Property {property.Name} is required. At row: {rowIndex + 1}");
    }

    public static bool? ThrowIfPropertyRequired(
        this bool? value,
        MemberInfo property,
        int rowIndex)
    {
        if (property.GetRequiredAttribute() == null || value != null) return value;
        throw new ValidationException($"Property {property.Name} is required. At row: {rowIndex + 1}");
    }

    public static Enum? ThrowIfPropertyRequired(
        this Enum? value,
        MemberInfo property,
        int rowIndex)
    {
        if (property.GetRequiredAttribute() == null || value != null) return value;
        throw new ValidationException($"Property {property.Name} is required. At row: {rowIndex + 1}");
    }

    public static RequiredAttribute? GetRequiredAttribute(this MemberInfo property)
    {
        var requiredAttribute = property.GetCustomAttribute<RequiredAttribute>();
        return requiredAttribute ?? null;
    }
}