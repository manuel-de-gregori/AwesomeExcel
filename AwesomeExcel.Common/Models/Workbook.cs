namespace AwesomeExcel.Common.Models;

/// <summary>
/// Represents an Excel Workbook.
/// </summary>
public class Workbook
{
    /// <summary>
    /// Sheets of the workbook.
    /// </summary>
    public IList<Sheet> Sheets { get; set; }

    /// <summary>
    /// File type.
    /// </summary>
    public FileType FileType { get; set; }
}
 