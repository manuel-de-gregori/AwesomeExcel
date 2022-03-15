using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.FluentCustomization;

public static class SheetCustomizationHeaderExtension
{
    public static SheetCustomization SetHeaderHorizontalAlignment(this SheetCustomization sheetInfo, HorizontalAlignment horizontalAlignment)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetHorizontalAlignment(horizontalAlignment);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderVerticalAlignment(this SheetCustomization sheetInfo, VerticalAlignment verticalAlignment)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetVerticalAlignment(verticalAlignment);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderBorderTopColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetBorderTopColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderBorderBottomColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetBorderBottomColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderBorderLeftColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetBorderLeftColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderBorderRightColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetBorderRightColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderFillForegroundColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetFillForegroundColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderFontName(this SheetCustomization sheetInfo, string name)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);
        InitializeHeaderFontStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetFontName(name);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderFontColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);
        InitializeHeaderFontStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetFontColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderFontHeightInPoints(this SheetCustomization sheetInfo, short height)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);
        InitializeHeaderFontStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetFontHeightInPoints(height);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderFontBold(this SheetCustomization sheetInfo, bool isBold)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);
        InitializeHeaderFontStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetFontBold(isBold);
        return sheetInfo;
    }

    public static SheetCustomization SetHeaderDateTimeFormat(this SheetCustomization sheetInfo, string format)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeHeaderStyle(sheetInfo);
        InitializeHeaderFontStyle(sheetInfo);

        sheetInfo.HeaderStyle.SetDateTimeFormat(format);
        return sheetInfo;
    }

    private static void InitializeHeaderStyle(SheetCustomization sheetInfo)
    {
        if (sheetInfo.HeaderStyle is null)
        {
            sheetInfo.HeaderStyle = new();
        }
    }

    private static void InitializeHeaderFontStyle(SheetCustomization sheetInfo)
    {
        if (sheetInfo.HeaderStyle.FontStyle is null)
        {
            sheetInfo.HeaderStyle.FontStyle = new();
        }
    }
}