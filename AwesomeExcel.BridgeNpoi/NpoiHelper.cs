using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class NpoiHelper
{
    public void SetCellValue(_NPOI.ICell cell, _Excel.ColumnType columnType, object value)
    {
        string valueStr = (value ?? "").ToString();

        if (columnType == _Excel.ColumnType.String)
        {
            cell.SetCellValue(valueStr);
        }
        else if (TryParseNumeric(columnType, value, valueStr, out double number))
        {
            cell.SetCellValue(number);
        }
        else if (TryParseDateTime(columnType, value, valueStr, out DateTime dt))
        {
            cell.SetCellValue(dt);
        }
        else
        {
            cell.SetCellValue(valueStr);
        }
    }

    public _NPOI.CellType GetCellType(_Excel.ColumnType columnType) => columnType switch
    {
        _Excel.ColumnType.DateTime => _NPOI.CellType.Numeric,
        _Excel.ColumnType.Numeric => _NPOI.CellType.Numeric,
        _Excel.ColumnType.String => _NPOI.CellType.String,
        _ => _NPOI.CellType.String,
    };

    private bool TryParseNumeric(_Excel.ColumnType columnType, object value, string valueStr, out double number)
    {
        if (columnType != _Excel.ColumnType.Numeric)
        {
            number = 0;
            return false;
        }

        switch (value)
        {
            case bool boo:
                number = Convert.ToDouble(boo);
                return true;
            case byte b:
                number = b;
                return true;
            case short s:
                number = s;
                return true;
            case ushort us:
                number = us;
                return true;
            case int i:
                number = i;
                return true;
            case uint ui:
                number = ui;
                return true;
            case long l:
                number = l;
                return true;
            case ulong ul:
                number = ul;
                return true;
            case float f:
                number = f;
                return true;
            case double d:
                number = d;
                return true;
            case decimal dec:
                number = Convert.ToDouble(dec);
                return true;

            default:
                return double.TryParse(valueStr, out number);
        }
    }

    private bool TryParseDateTime(_Excel.ColumnType columnType, object value, string valueStr, out DateTime dt)
    {
        if (columnType != _Excel.ColumnType.DateTime)
        {
            dt = default;
            return false;
        }

        if (value is DateTime _datetime)
        {
            dt = _datetime;
            return true;
        }
        else if (value is DateOnly date)
        {
            dt = date.ToDateTime(new TimeOnly(0, 0, 0, 0));
            return true;
        }
        else if (value is TimeOnly time)
        {
            DateTime minDateTime = DateTime.MinValue;
            dt = new DateTime(minDateTime.Year, minDateTime.Month, minDateTime.Day, time.Hour, time.Minute, time.Second);
            return true;
        }
        else
        {
            return DateTime.TryParse(valueStr, out dt);
        }
    }
}