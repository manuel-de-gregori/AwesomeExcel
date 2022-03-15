using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Services;
using AwesomeExcel.Generator.Services;

namespace AwesomeExcel.Generator;

/// <summary>
/// A base class for an Excel file generator.
/// </summary>
/// <typeparam name="TWorkbook"></typeparam>
public abstract class FileGenerator<TWorkbook>
{
    private readonly WorkbookFactory workbookFactory = new();
    private readonly SheetFactory sheetFactory = new();

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet">The type of the worksheet rows.</typeparam>
    /// <param name="rows">The rows of the sheet.</param>
    /// <param name="action">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet>(IEnumerable<TSheet> rows, Action<SingleSheetCustomizationService<TSheet>> action)
    {
        SingleSheetCustomizationService<TSheet> sps = new();
        action(sps);

        Sheet sheet = sheetFactory.Create(rows, sps.Sheet, sps.ccs.GetCustomizedColumn());

        Workbook excelWorkbook = workbookFactory.Create(new Sheet[1] { sheet }, sps.Workbook);
        return GetStream(excelWorkbook);
    }

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet1">The type of the rows of the first sheet.</typeparam>
    /// <typeparam name="TSheet2">The type of the rows of the second sheet.</typeparam>
    /// <param name="rowsSheet1">The rows of the first sheet.</param>
    /// <param name="rowsSheet2">The rows of the second sheet.</param>
    /// <param name="action">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet1, TSheet2>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, Action<MultipleSheetsCustomizationService<TSheet1, TSheet2>> action)
    {
        MultipleSheetsCustomizationService<TSheet1, TSheet2> msps = new();
        action(msps);

        Sheet sheet1 = sheetFactory.Create(rowsSheet1, msps.Sheets.Sheet1, msps.GetColumnCustomizationService(msps.Sheets.Sheet1).GetCustomizedColumn());
        Sheet sheet2 = sheetFactory.Create(rowsSheet2, msps.Sheets.Sheet2, msps.GetColumnCustomizationService(msps.Sheets.Sheet2).GetCustomizedColumn());

        Workbook excelWorkbook = workbookFactory.Create(new Sheet[] { sheet1, sheet2 }, msps.Workbook);
        return GetStream(excelWorkbook);
    }

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet1">The type of the rows of the first sheet.</typeparam>
    /// <typeparam name="TSheet2">The type of the rows of the second sheet.</typeparam>
    /// <typeparam name="TSheet3">The type of the rows of the third sheet.</typeparam>
    /// <param name="rowsSheet1">The rows of the first sheet.</param>
    /// <param name="rowsSheet2">The rows of the second sheet.</param>
    /// <param name="rowsSheet3">The rows of the third sheet.</param>
    /// <param name="action">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet1, TSheet2, TSheet3>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, IEnumerable<TSheet3> rowsSheet3, Action<MultipleSheetsCustomizationService<TSheet1, TSheet2, TSheet3>> action)
    {
        MultipleSheetsCustomizationService<TSheet1, TSheet2, TSheet3> msps = new();
        action(msps);

        Sheet sheet1 = sheetFactory.Create(rowsSheet1, msps.Sheets.Item1, msps.GetColumnCustomizationService(msps.Sheets.Item1).GetCustomizedColumn());
        Sheet sheet2 = sheetFactory.Create(rowsSheet2, msps.Sheets.Item2, msps.GetColumnCustomizationService(msps.Sheets.Item2).GetCustomizedColumn());
        Sheet sheet3 = sheetFactory.Create(rowsSheet3, msps.Sheets.Item3, msps.GetColumnCustomizationService(msps.Sheets.Item3).GetCustomizedColumn());

        Workbook excelWorkbook = workbookFactory.Create(new List<Sheet> { sheet1, sheet2, sheet3 }, msps.Workbook);
        return GetStream(excelWorkbook);
    }

    private MemoryStream GetStream(Workbook excelWorkbook)
    {
        TWorkbook npoiWorkbook = Convert(excelWorkbook);
        MemoryStream stream = Write(npoiWorkbook);
        return stream;
    }

    /// <summary>
    /// Convert a common Workbook to a library specific implementation of a Workbook.
    /// </summary>
    /// <param name="workbook">The workbook to be converted.</param>
    /// <returns>The converted workbook.</returns>
    protected abstract TWorkbook Convert(Workbook workbook);

    /// <summary>
    /// Write a specific implementation of a Workbook to a stream.
    /// </summary>
    /// <param name="workbook">The workbook to be written.</param>
    /// <returns>The stream the Excel file is written into.</returns>
    protected abstract MemoryStream Write(TWorkbook workbook);
}


