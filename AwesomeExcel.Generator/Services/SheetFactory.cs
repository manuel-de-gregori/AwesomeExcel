using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("AwesomeExcel.Generator.UnitTests")]

namespace AwesomeExcel.Generator.Services;

/// <summary>
/// Defines a factory of objects of type Sheet.
/// </summary>
internal class SheetFactory
{
    /// <summary>
    /// Create a new Sheet object.
    /// </summary>
    /// <typeparam name="TSheet">The type of the rows of the sheet.</typeparam>
    /// <param name="rows">The rows of the sheet.</param>
    /// <param name="si">Additional informations used to customize the sheet.</param>
    /// <param name="columnCustomizationService">Additional informations used to customize the columns of the sheet.</param>
    /// <returns>An Excel Sheet with the given rows and customizations.</returns>
    /// <exception cref="ArgumentNullException">rows is null</exception>
    /// <exception cref="InvalidOperationException">rows contains null elements</exception>
    public Sheet Create<TSheet>(IEnumerable<TSheet> rows, SheetCustomization si, Dictionary<PropertyInfo, ColumnCustomization> ci)
    {
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }

        if (rows.Any(r => r is null))
        {
            throw new InvalidOperationException(nameof(rows));
        }

        PropertyInfo[] properties = typeof(TSheet).GetProperties();
        List<Row> sheetRows = GetRows(rows, properties);
        List<Column> sheetColumns = GetColumns(properties, ci);

        return new Sheet
        {
            Name = si?.Name,
            Rows = sheetRows,
            Columns = sheetColumns,
            HasHeader = si?.HasHeader ?? false,
            Style = si?.Style,
            HeaderStyle = si?.HeaderStyle,
            IsProtected = si?.IsProtected ?? false
        };
    }

    private List<Row> GetRows(IEnumerable rows, IEnumerable<PropertyInfo> properties)
    {
        List<Row> sheetRows = new();

        foreach (object row in rows)
        {
            Row excelRow = GetRow(row, properties);
            sheetRows.Add(excelRow);
        }

        return sheetRows;
    }

    private Row GetRow(object row, IEnumerable<PropertyInfo> properties)
    {
        IList<Cell> cells = properties
            .Select(pi => GetCell(row, pi))
            .ToList();

        return new Row
        {
            Cells = cells
        };
    }

    private Cell GetCell(object row, PropertyInfo pi)
    {
        object pValue = pi.GetValue(row, null);
        return new Cell
        {
            Value = pValue
        };
    }

    private List<Column> GetColumns(IEnumerable<PropertyInfo> properties, Dictionary<PropertyInfo, ColumnCustomization> ci)
    {
        List<Column> excelColumns = properties
            .Where(pi => ci is null || !ci.TryGetValue(pi, out ColumnCustomization value) || value.Excluded == false)
            .Select(pi =>
            {
                ColumnCustomization columnInfo = null;
                ci?.TryGetValue(pi, out columnInfo);

                string columnName = columnInfo?.Name ?? pi.Name;
                Style columnStyle = columnInfo?.Style;
                return GetColumn(columnName, pi.PropertyType, columnStyle);
            })
            .ToList();

        return excelColumns;
    }

    private Column GetColumn(string name, Type type, Style style)
    {
        return new Column()
        {
            Name = name,
            ColumnType = GetColumnType(type),
            Style = style
        };
    }

    private ColumnType GetColumnType(Type type)
    {
        Dictionary<Type, ColumnType> conversionTable = new()
        {
            { typeof(char), ColumnType.String },
            { typeof(string), ColumnType.String },
            { typeof(DateTime), ColumnType.DateTime },
            { typeof(DateTimeOffset), ColumnType.DateTime },
            { typeof(DateOnly), ColumnType.DateTime },
            { typeof(TimeOnly), ColumnType.DateTime },
            { typeof(bool), ColumnType.Numeric },
            { typeof(byte), ColumnType.Numeric },
            { typeof(short), ColumnType.Numeric },
            { typeof(ushort), ColumnType.Numeric },
            { typeof(int), ColumnType.Numeric },
            { typeof(uint), ColumnType.Numeric },
            { typeof(long), ColumnType.Numeric },
            { typeof(ulong), ColumnType.Numeric },
            { typeof(float), ColumnType.Numeric },
            { typeof(double), ColumnType.Numeric },
            { typeof(decimal), ColumnType.Numeric }
        };

        type = Nullable.GetUnderlyingType(type) ?? type;

        if (conversionTable.TryGetValue(type, out ColumnType value))
        {
            return value;
        }
        else if (type.IsClass)
        {
            throw new Exception();
        }
        else if (type.IsEnum)
        {
            return ColumnType.String;
        }
        else if (type.IsValueType)
        {
            return ColumnType.String;
        }
        else
        {
            throw new NotSupportedException();
        }
    }
}
