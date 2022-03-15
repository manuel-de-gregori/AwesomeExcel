using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public class SingleSheetCustomizationService<T>
{
    public readonly ColumnCustomizationService ccs = new();
    internal readonly CellsCustomizationService<T> cells = new();

    public WorkbookCustomization Workbook { get; } = new();
    public SheetCustomization Sheet { get; } = new();
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
}
