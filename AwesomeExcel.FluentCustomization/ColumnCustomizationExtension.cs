using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.FluentCustomization;

public static class ColumnCustomizationExtension
{
    public static ColumnCustomization SetName(this ColumnCustomization columnInfo, string name)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        columnInfo.Name = name;
        return columnInfo;
    }

    public static ColumnCustomization SetStyle(this ColumnCustomization columnInfo, Action<Style> fn)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);
        InitializeFontStyle(columnInfo);

        fn(columnInfo.Style);
        return columnInfo;
    }

    public static ColumnCustomization SetHorizontalAlignment(this ColumnCustomization columnInfo, HorizontalAlignment horizontalAlignment)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetHorizontalAlignment(horizontalAlignment);
        return columnInfo;
    }

    public static ColumnCustomization SetVerticalAlignment(this ColumnCustomization columnInfo, VerticalAlignment verticalAlignment)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetVerticalAlignment(verticalAlignment);
        return columnInfo;
    }

    public static ColumnCustomization SetBorderTopColor(this ColumnCustomization columnInfo, Color color)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetBorderTopColor(color);
        return columnInfo;
    }

    public static ColumnCustomization SetBorderBottomColor(this ColumnCustomization columnInfo, Color color)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetBorderBottomColor(color);
        return columnInfo;
    }

    public static ColumnCustomization SetBorderLeftColor(this ColumnCustomization columnInfo, Color color)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetBorderLeftColor(color);
        return columnInfo;
    }

    public static ColumnCustomization SetBorderRightColor(this ColumnCustomization columnInfo, Color color)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetBorderRightColor(color);
        return columnInfo;
    }

    public static ColumnCustomization SetFillForegroundColor(this ColumnCustomization columnInfo, Color color)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetFillForegroundColor(color);
        return columnInfo;
    }

    public static ColumnCustomization SetFontName(this ColumnCustomization columnInfo, string name)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);
        InitializeFontStyle(columnInfo);

        columnInfo.Style.SetFontName(name);
        return columnInfo;
    }

    public static ColumnCustomization SetFontColor(this ColumnCustomization columnInfo, Color color)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);
        InitializeFontStyle(columnInfo);

        columnInfo.Style.SetFontColor(color);
        return columnInfo;
    }

    public static ColumnCustomization SetFontHeightInPoints(this ColumnCustomization columnInfo, short height)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);
        InitializeFontStyle(columnInfo);

        columnInfo.Style.SetFontHeightInPoints(height);
        return columnInfo;
    }

    public static ColumnCustomization SetFontBold(this ColumnCustomization columnInfo, bool isBold)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);
        InitializeFontStyle(columnInfo);

        columnInfo.Style.SetFontBold(isBold);
        return columnInfo;
    }

    public static ColumnCustomization SetDateTimeFormat(this ColumnCustomization columnInfo, string format)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        InitializeStyle(columnInfo);

        columnInfo.Style.SetDateTimeFormat(format);
        return columnInfo;
    }

    private static void InitializeStyle(ColumnCustomization columnInfo)
    {
        if (columnInfo.Style is null)
        {
            columnInfo.Style = new();
        }
    }

    private static void InitializeFontStyle(ColumnCustomization columnInfo)
    {
        if (columnInfo.Style.FontStyle is null)
        {
            columnInfo.Style.FontStyle = new();
        }
    }

    public static ColumnCustomization Exclude(this ColumnCustomization columnInfo)
    {
        if (columnInfo is null)
        {
            throw new ArgumentNullException(nameof(columnInfo));
        }

        columnInfo.Excluded = true;
        return columnInfo;
    }
}
