using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;

namespace AwesomeExcel.Customization.Services;

public interface ISheetsCustomizer
{
    WorkbookCustomization Workbook { get; }
    ColumnCustomization Column<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector);
    CellCustomization<TProperty> Cells<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector);
}

public interface ISheetsCustomizer<TSheet1, TSheet2> : ISheetsCustomizer
{
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
}

public interface ISheetsCustomizer<TSheet1, TSheet2, TSheet3> : ISheetsCustomizer
{
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
    SheetCustomization<TSheet3> Sheet3 { get; }
}

public interface ISheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4> : ISheetsCustomizer
{
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
    SheetCustomization<TSheet3> Sheet3 { get; }
    SheetCustomization<TSheet4> Sheet4 { get; }
}

public interface ISheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5> : ISheetsCustomizer
{
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
    SheetCustomization<TSheet3> Sheet3 { get; }
    SheetCustomization<TSheet4> Sheet4 { get; }
    SheetCustomization<TSheet5> Sheet5 { get; }
}
