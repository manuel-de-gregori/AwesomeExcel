namespace AwesomeExcel.Common.Models;

/// <summary>
/// Represents a sheet of a workbook.
/// </summary>
public class Sheet
{
    /// <summary>
    /// Gets or sets the name of the sheet.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Determines whether the sheet is read only.
    /// </summary>
    public bool IsProtected { get; set; }

    /// <summary>
    /// Columns of the sheet.
    /// </summary>
    public IList<Column> Columns { get; set; }

    /// <summary>
    /// Rows of the sheet.
    /// </summary>
    public IList<Row> Rows { get; set; }

    /// <summary>
    /// Gets or sets the style of the sheet.
    /// </summary>
    public SheetStyle Style { get; set; }

    /// <summary>
    /// Determines whether the sheet has a header row.
    /// </summary>
    public bool HasHeader { get; set; }

    /// <summary>
    /// Gets or sets the style of the header.
    /// </summary>
    public Style HeaderStyle { get; set; }
}

public class Sheet<T> : Sheet { }