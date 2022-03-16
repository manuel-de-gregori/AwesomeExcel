using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public class SingleSheetCustomizer<T> : ISheetCustomizer<T>
{
    private readonly ColumnsCustomizer ccs = new();
    private readonly CellsCustomizer<T> cells = new();

    public WorkbookCustomization Workbook { get; } = new();
    public SheetCustomization<T> Sheet { get; } = new();
    public ColumnCustomization GetColumn<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        var x = ccs.GetOrCreateColumn(pi);
        return x;
    }
    public CellCustomization GetCells<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        return cells.GetOrCreateCells(pi);
    }

    public IReadOnlyDictionary<PropertyInfo, ColumnCustomization> GetCustomizedColumns()
    {
        IReadOnlyDictionary<PropertyInfo, ColumnCustomization> x = ccs.GetCustomizedColumn();
        return x;
    }
}
