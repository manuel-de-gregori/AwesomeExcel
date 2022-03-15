using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;


namespace AwesomeExcel.BridgeNpoi;

internal interface IStyleConverter
{
    public _NPOI.ICellStyle Convert(_Excel.Style style);

    public _NPOI.IFont Convert(_Excel.FontStyle style);
}
