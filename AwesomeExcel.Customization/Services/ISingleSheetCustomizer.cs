using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;

namespace AwesomeExcel.Customization.Services;

public interface ISingleSheetCustomizer<T>
{
    SheetCustomization<T> Sheet { get; }
    WorkbookCustomization Workbook { get; }

    CellCustomization GetCells<TProperty>(Expression<Func<T, TProperty>> selector);
    ColumnCustomization GetColumn<TProperty>(Expression<Func<T, TProperty>> selector);
}
