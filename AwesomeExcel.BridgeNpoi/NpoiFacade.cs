using NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class NpoiFacade
{
    public ICell CreateEmptyCell(IRow npoiRow, int columnIndex)
    {
        ICell cell = npoiRow.CreateCell(columnIndex, CellType.String);
        string blank = null;
        cell.SetCellValue(blank);
        return cell;
    }

    public ICell CreateCell(IRow row, int columnIndex, CellType cellType, ICellStyle cellStyle)
    {
        ICell cell = row.CreateCell(columnIndex, cellType);
        cell.CellStyle = cellStyle;
        return cell;
    }

    public ISheet CreateSheet(IWorkbook workbook, string name)
    {
        return string.IsNullOrWhiteSpace(name)
            ? workbook.CreateSheet()
            : workbook.CreateSheet(name);
    }

    public void AutoSizeColumns(ISheet sheet, int columnsCount)
    {
        for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
        {
            sheet.AutoSizeColumn(columnIndex);

            // After the autosize the width is still small, it's better to increase it a little bit more
            int widthAfterAutoSize = sheet.GetColumnWidth(columnIndex);
            int width = (int)(widthAfterAutoSize * 1.5);
            sheet.SetColumnWidth(columnIndex, width);
        }
    }

    public void CreateHeaderRow(ISheet sheet, ICellStyle headerStyle, IEnumerable<string> columnsName)
    {
        IRow headerRow = sheet.CreateRow(0);

        // The header needs to be higher than normal rows
        headerRow.HeightInPoints *= 1.5f;

        int columnIndex = 0;
        foreach (string columnName in columnsName)
        {
            ICell cell = CreateCell(headerRow, columnIndex, CellType.String, headerStyle);
            cell.SetCellValue(columnName);
            columnIndex++;
        }
    }

    public IRow CreateRow(ISheet sheet, int rowNumber, ICellStyle rowStyle)
    {
        IRow row = sheet.CreateRow(rowNumber);
        row.RowStyle = rowStyle;
        return row;
    }
}
