using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;

namespace AwesomeExcel.Customization.Services;

public interface ISheetCustomizer<TSheet>
{
    SheetCustomization<TSheet> Sheet { get; }
    WorkbookCustomization Workbook { get; }

    CellCustomization<TProperty> Cells<TProperty>(Expression<Func<TSheet, TProperty>> selector);
    ColumnCustomization Column<TProperty>(Expression<Func<TSheet, TProperty>> selector);
}
