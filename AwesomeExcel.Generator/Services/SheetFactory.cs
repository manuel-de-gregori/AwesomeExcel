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
    /// <param name="sheetCustomization">Additional informations used to customize the sheet.</param>
    /// <param name="columnCustomizationService">Additional informations used to customize the columns of the sheet.</param>
    /// <returns>An Excel Sheet with the given rows and customizations.</returns>
    /// <exception cref="ArgumentNullException">rows is null</exception>
    /// <exception cref="InvalidOperationException">rows contains null elements</exception>
    public Sheet Create<TSheet>(
        IEnumerable<TSheet> rows,
        SheetCustomization<TSheet> sheetCustomization,
        IReadOnlyDictionary<PropertyInfo, ColumnCustomization> columnsCustomization,
        IReadOnlyDictionary<PropertyInfo, CellCustomization> cellsCustomization)
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
        List<Row> sheetRows = GetRows(rows, properties, cellsCustomization);
        List<Column> sheetColumns = GetColumns(properties, columnsCustomization);

        return new Sheet
        {
            Name = sheetCustomization?.Name,
            Rows = sheetRows,
            Columns = sheetColumns,
            HasHeader = sheetCustomization?.HasHeader ?? false,
            Style = sheetCustomization?.Style,
            HeaderStyle = sheetCustomization?.HeaderStyle,
            IsProtected = sheetCustomization?.IsProtected ?? false
        };
    }

    private List<Row> GetRows(IEnumerable rows, IEnumerable<PropertyInfo> properties, IReadOnlyDictionary<PropertyInfo, CellCustomization> cellsCustomization)
    {
        List<Row> sheetRows = new();

        foreach (object row in rows)
        {
            Row excelRow = GetRow(row, properties, cellsCustomization);
            sheetRows.Add(excelRow);
        }

        return sheetRows;
    }

    private Row GetRow(object row, IEnumerable<PropertyInfo> properties, IReadOnlyDictionary<PropertyInfo, CellCustomization> cc)
    {
        IList<Cell> cells = properties
            .Select(pi =>
            {
                object pValue = pi.GetValue(row, null);

                CellCustomization cellCustomization = null;
                cc?.TryGetValue(pi, out cellCustomization);

                if (cellCustomization is null)
                {
                    return new Cell
                    {
                        Value = pValue,
                        Style = null
                    };
                }
                else
                {
                    Style s = GetCellStyle(cellCustomization, pValue);
                    return new Cell
                    {
                        Value = pValue,
                        Style = s
                    };
                }
            })
            .ToList();

        return new Row
        {
            Cells = cells
        };
    }

    private List<Column> GetColumns(IEnumerable<PropertyInfo> properties, IReadOnlyDictionary<PropertyInfo, ColumnCustomization> ci)
    {
        List<Column> excelColumns = properties
            .Where(pi => ci is null || !ci.TryGetValue(pi, out ColumnCustomization value) || value.Excluded == false)
            .Select(pi =>
            {
                ColumnCustomization columnInfo = null;
                ci?.TryGetValue(pi, out columnInfo);

                string columnName = columnInfo?.Name ?? pi.Name;
                Style columnStyle = columnInfo?.Style;

                var excelColumnType = GetColumnType(pi.PropertyType);

                return new Column()
                {
                    Name = columnName,
                    ColumnType = excelColumnType,
                    Style = columnStyle
                };
            })
            .ToList();

        return excelColumns;
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

    private Style GetCellStyle(CellCustomization customization, object value)
    {
        StyleCustomization sc = (StyleCustomization)customization.GetType().GetProperty("Style").GetValue(customization);

        if (sc is null)
            return null;

        Type scType = sc.GetType();
        StyleCustomization<object> _;

        var borderTopColorFn = (Func<object, Color?>)scType.GetProperty(nameof(_.BorderTopColor)).GetValue(sc);
        var borderBottomColorFn = (Func<object, Color?>)scType.GetProperty(nameof(_.BorderBottomColor)).GetValue(sc);
        var borderLeftColorFn = (Func<object, Color?>)scType.GetProperty(nameof(_.BorderLeftColor)).GetValue(sc);
        var borderRightColorFn = (Func<object, Color?>)scType.GetProperty(nameof(_.BorderRightColor)).GetValue(sc);
        var fillForegroundColorFn = (Delegate)scType.GetProperty(nameof(_.FillForegroundColor)).GetValue(sc);
        var fillPatternFn = (Func<object, FillPattern?>)scType.GetProperty(nameof(_.FillPattern)).GetValue(sc);
        var dateTimeFormatFn = (Func<object, string>)scType.GetProperty(nameof(_.DateTimeFormat)).GetValue(sc);
        var horizontalAlignmentFn = (Func<object, HorizontalAlignment?>)scType.GetProperty(nameof(_.HorizontalAlignment)).GetValue(sc);
        var verticalAlignmentFn = (Func<object, VerticalAlignment?>)scType.GetProperty(nameof(_.VerticalAlignment)).GetValue(sc);

        var fontStyleCustomization = (FontStyleCustomization)scType.GetProperty(nameof(_.FontStyle)).GetValue(sc);

        FontStyle fontStyle = null;

        if (fontStyleCustomization is not null)
        {
            Type fscType = fontStyleCustomization.GetType();
            FontStyleCustomization<object> __;

            var colorFn = (Func<object, Color?>)fscType.GetProperty(nameof(__.Color)).GetValue(fontStyleCustomization);
            var heightInPointsFn = (Func<object, short?>)fscType.GetProperty(nameof(__.HeightInPoints)).GetValue(fontStyleCustomization);
            var isBoldFn = (Func<object, bool?>)fscType.GetProperty(nameof(__.IsBold)).GetValue(fontStyleCustomization);
            var nameFn = (Func<object, string>)fscType.GetProperty(nameof(__.Name)).GetValue(fontStyleCustomization);

            fontStyle = new FontStyle
            {
                Color = colorFn?.Invoke(value),
                HeightInPoints = heightInPointsFn?.Invoke(value),
                IsBold = isBoldFn?.Invoke(value),
                Name = nameFn?.Invoke(value)
            };
        }

        return new Style
        {
            BorderTopColor = borderTopColorFn?.Invoke(value),
            BorderBottomColor = borderBottomColorFn?.Invoke(value),
            BorderLeftColor = borderLeftColorFn?.Invoke(value),
            BorderRightColor = borderRightColorFn?.Invoke(value),
            FillForegroundColor = (Color?)fillForegroundColorFn?.Method.Invoke(fillForegroundColorFn.Target, new[] { value }),
            FillPattern = fillPatternFn?.Invoke(value),
            DateTimeFormat = dateTimeFormatFn?.Invoke(value),
            HorizontalAlignment = horizontalAlignmentFn?.Invoke(value),
            VerticalAlignment = verticalAlignmentFn?.Invoke(value),
            FontStyle = fontStyle
        };
    }
}
