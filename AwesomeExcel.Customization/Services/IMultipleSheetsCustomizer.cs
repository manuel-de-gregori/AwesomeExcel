using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.Customization.Services;

public interface IMultipleSheetsCustomizer<TSheet1, TSheet2>
{
    WorkbookCustomization Workbook { get; }
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
}

public interface IMultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3>
{
    WorkbookCustomization Workbook { get; }
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
    SheetCustomization<TSheet3> Sheet3 { get; }
}

public interface IMultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4>
{
    WorkbookCustomization Workbook { get; }
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
    SheetCustomization<TSheet3> Sheet3 { get; }
    SheetCustomization<TSheet4> Sheet4 { get; }
}

public interface IMultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5>
{
    WorkbookCustomization Workbook { get; }
    SheetCustomization<TSheet1> Sheet1 { get; }
    SheetCustomization<TSheet2> Sheet2 { get; }
    SheetCustomization<TSheet3> Sheet3 { get; }
    SheetCustomization<TSheet4> Sheet4 { get; }
    SheetCustomization<TSheet5> Sheet5 { get; }
}
