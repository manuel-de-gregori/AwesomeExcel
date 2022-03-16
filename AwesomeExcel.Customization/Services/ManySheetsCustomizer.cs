using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public abstract class ManySheetsCustomizer : IManySheetsCustomizer
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

    public CellCustomization<TProperty> GetCells<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        return null;
    }

    public IReadOnlyDictionary<PropertyInfo, ColumnCustomization> GetCustomizedColumns<T>(SheetCustomization<T> sheet)
    {
        IReadOnlyDictionary<PropertyInfo, ColumnCustomization> ccs = dict[sheet].GetCustomizedColumn();
        return ccs;
    }
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2> : ManySheetsCustomizer, ISheetsCustomizer<TSheet1, TSheet2>
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3> : ManySheetsCustomizer, ISheetsCustomizer<TSheet1, TSheet2, TSheet3>
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4> : ManySheetsCustomizer, ISheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4>
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
    public SheetCustomization<TSheet4> Sheet4 { get; } = new();
}

public class MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5> : ManySheetsCustomizer, ISheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5>
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
    public SheetCustomization<TSheet4> Sheet4 { get; } = new();
    public SheetCustomization<TSheet5> Sheet5 { get; } = new();
}
