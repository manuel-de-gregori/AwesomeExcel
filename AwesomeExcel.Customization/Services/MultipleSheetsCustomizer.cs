using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public abstract class MultipleSheetsCustomizer
{
    public WorkbookCustomization Workbook { get; } = new();

    private Dictionary<SheetCustomization, ColumnsCustomizer> dict = new();

    public ColumnCustomization GetColumn<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        if (!dict.ContainsKey(sheet))
        {
            dict.Add(sheet, new ColumnsCustomizer());
        }

        var customizer = dict[sheet];

        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        return customizer.GetOrCreateColumn(pi);
    }

    public Cell<TProperty> GetCells<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        return null;
    }

    public ColumnsCustomizer GetColumnCustomizationService<T>(SheetCustomization<T> sheet)
    {
        ColumnsCustomizer ccs = dict[sheet];
        return ccs;
    }
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2> : MultipleSheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3> : MultipleSheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4> : MultipleSheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
    public SheetCustomization<TSheet3> Sheet4 { get; } = new();
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5> : MultipleSheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
    public SheetCustomization<TSheet3> Sheet4 { get; } = new();
    public SheetCustomization<TSheet3> Sheet5 { get; } = new();
}
