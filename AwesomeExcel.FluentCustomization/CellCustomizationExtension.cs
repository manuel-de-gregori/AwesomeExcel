using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.FluentCustomization;

public static class CellCustomizationExtension
{
    public static CellCustomization<T> SetHorizontalAlignment<T>(this CellCustomization<T> cellInfo, Func<T, HorizontalAlignment?> horizontalAlignment)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.HorizontalAlignment = horizontalAlignment;
        return cellInfo;
    }

    public static CellCustomization<T> SetVerticalAlignment<T>(this CellCustomization<T> cellInfo, Func<T, VerticalAlignment?> verticalAlignment)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.VerticalAlignment = verticalAlignment;
        return cellInfo;
    }

    public static CellCustomization<T> SetBorderTopColor<T>(this CellCustomization<T> cellInfo, Func<T, Color?> color)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.BorderTopColor = color;
        return cellInfo;
    }

    public static CellCustomization<T> SetBorderBottomColor<T>(this CellCustomization<T> cellInfo, Func<T, Color?> color)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.BorderBottomColor = color;
        return cellInfo;
    }

    public static CellCustomization<T> SetBorderLeftColor<T>(this CellCustomization<T> cellInfo, Func<T, Color?> color)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.BorderLeftColor = color;
        return cellInfo;
    }

    public static CellCustomization<T> SetBorderRightColor<T>(this CellCustomization<T> cellInfo, Func<T, Color?> color)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.BorderRightColor = color;
        return cellInfo;
    }

    public static CellCustomization<T> SetFillForegroundColor<T>(this CellCustomization<T> cellInfo, Func<T, Color?> color)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.FillForegroundColor = color;
        return cellInfo;
    }

    public static CellCustomization<T> SetFontName<T>(this CellCustomization<T> cellInfo, Func<T, string> name)
    {
        InitializeStyle(cellInfo);
        InitializeFontStyle(cellInfo);

        cellInfo.Style.FontStyle.Name = name;
        return cellInfo;
    }

    public static CellCustomization<T> SetFontColor<T>(this CellCustomization<T> cellInfo, Func<T, Color?> color)
    {
        InitializeStyle(cellInfo);
        InitializeFontStyle(cellInfo);

        cellInfo.Style.FontStyle.Color = color;
        return cellInfo;
    }

    public static CellCustomization<T> SetFontHeightInPoints<T>(this CellCustomization<T> cellInfo, Func<T, short?> height)
    {
        InitializeStyle(cellInfo);
        InitializeFontStyle(cellInfo);

        cellInfo.Style.FontStyle.HeightInPoints = height;
        return cellInfo;
    }

    public static CellCustomization<T> SetFontBold<T>(this CellCustomization<T> cellInfo, Func<T, bool?> isBold)
    {
        InitializeStyle(cellInfo);
        InitializeFontStyle(cellInfo);

        cellInfo.Style.FontStyle.IsBold = isBold;
        return cellInfo;
    }

    public static CellCustomization<T> SetDateTimeFormat<T>(this CellCustomization<T> cellInfo, Func<T, string> format)
    {
        InitializeStyle(cellInfo);

        cellInfo.Style.DateTimeFormat = format;
        return cellInfo;
    }

    private static void InitializeStyle<T>(CellCustomization<T> cellInfo)
    {
        if (cellInfo.Style is null)
        {
            cellInfo.Style = new();
        }
    }

    private static void InitializeFontStyle<T>(CellCustomization<T> cellInfo)
    {
        if (cellInfo.Style.FontStyle is null)
        {
            cellInfo.Style.FontStyle = new();
        }
    }
}
