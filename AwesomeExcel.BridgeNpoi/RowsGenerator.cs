using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class RowsGenerator
{
    private readonly NpoiFacade npoiFacade = new();
    private readonly Common.Services.StylesMerger stylesMerger = new();
    private readonly NpoiHelper npoiHelper = new();

    private readonly _NPOI.ISheet npoiSheet;
    private readonly _Excel.Sheet excelSheet;
    private readonly IStyleConverter styleConverter;

    private _Excel.Style colorBandingEvenRows;
    private _Excel.Style colorbandingOddRows;

    public RowsGenerator(_NPOI.ISheet npoiSheet, _Excel.Sheet excelSheet, IStyleConverter styleConverter)
    {
        this.npoiSheet = npoiSheet;
        this.excelSheet = excelSheet;
        this.styleConverter = styleConverter;
    }

    public void GenerateHeaderRow()
    {
        if (!excelSheet.HasHeader)
            return;

        _Excel.Style excelStyle = stylesMerger.Merge(excelSheet.Style, excelSheet.HeaderStyle);
        _NPOI.ICellStyle npoiStyle = styleConverter.Convert(excelStyle);
        IEnumerable<string> columns = excelSheet.Columns.Select(c => c.Name);

        npoiFacade.CreateHeaderRow(npoiSheet, npoiStyle, columns);
    }

    public void GenerateRows()
    {
        for (int rowIndex = 0; rowIndex < (excelSheet.Rows?.Count ?? 0); rowIndex++)
        {
            int rowNumber = rowIndex + (excelSheet.HasHeader ? 1 : 0);

            _Excel.Row row = excelSheet.Rows[rowIndex];
            _Excel.Style rowStyle = stylesMerger.Merge(excelSheet.Style, row?.Style);

            _NPOI.ICellStyle npoiStyle = styleConverter.Convert(rowStyle);
            _NPOI.IRow npoiRow = npoiFacade.CreateRow(npoiSheet, rowNumber, npoiStyle);

            GenerateCells(npoiRow, row, rowNumber);
        }
    }

    private _Excel.Style GetColorBandingStyle(int rowNumber)
    {
        if (excelSheet.Style?.ColorBanding is null)
            return null;

        bool isEven = rowNumber % 2 == 0;

        if (isEven && colorBandingEvenRows != null)
        {
            return colorBandingEvenRows;
        }
        else if (!isEven && colorbandingOddRows != null)
        {
            return colorbandingOddRows;
        }

        _Excel.ColorBanding colorBanding = excelSheet.Style?.ColorBanding;
        _Excel.Color color = isEven ? colorBanding.EvenRows : colorBanding.OddRows;

        var s = new _Excel.Style()
        {
            FillForegroundColor = color,
            FillPattern = _Excel.FillPattern.SolidForeground
        };
        if (isEven)
        {
            colorBandingEvenRows = s;
        }
        else
        {
            colorbandingOddRows = s;
        }
        return s;
    }

    private void GenerateCells(_NPOI.IRow npoiRow, _Excel.Row excelRow, int rowNumber)
    {
        for (int columnIndex = 0; columnIndex < (excelRow?.Cells?.Count ?? 0); columnIndex++)
        {
            _Excel.Column column;
            if (columnIndex < (excelSheet.Columns?.Count ?? 0))
            {
                column = excelSheet.Columns[columnIndex];
            }
            else
            {
                column = new()
                {
                    ColumnType = _Excel.ColumnType.String,
                    Name = null,
                    Style = null
                };
            }

            _Excel.Cell cell = excelRow.Cells[columnIndex];
            _Excel.Style colorBandingStyle = GetColorBandingStyle(rowNumber);
            _Excel.Style style = stylesMerger.Merge(excelSheet.Style, colorBandingStyle, column.Style, cell?.Style);

            _NPOI.CellType cellType = npoiHelper.GetCellType(column.ColumnType);
            _NPOI.ICellStyle npoiStyle = styleConverter.Convert(style);
            _NPOI.ICell npoiCell = npoiFacade.CreateCell(npoiRow, columnIndex, cellType, npoiStyle);
            npoiHelper.SetCellValue(npoiCell, column.ColumnType, cell?.Value);
        }
    }
}
