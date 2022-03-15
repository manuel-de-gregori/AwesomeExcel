using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.FluentCustomization;

public static class SheetCustomizationStyleExtension
{
    public static SheetCustomization SetHorizontalAlignment(this SheetCustomization sheetInfo, HorizontalAlignment horizontalAlignment)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetHorizontalAlignment(horizontalAlignment);
        return sheetInfo;
    }

    public static SheetCustomization SetVerticalAlignment(this SheetCustomization sheetInfo, VerticalAlignment verticalAlignment)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetVerticalAlignment(verticalAlignment);
        return sheetInfo;
    }

    public static SheetCustomization SetBorderTopColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetBorderTopColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetBorderBottomColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetBorderBottomColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetBorderLeftColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetBorderLeftColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetBorderRightColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetBorderRightColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetFillForegroundColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetFillForegroundColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetFontName(this SheetCustomization sheetInfo, string name)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);
        InitializeFontStyle(sheetInfo);

        sheetInfo.Style.SetFontName(name);
        return sheetInfo;
    }

    public static SheetCustomization SetFontColor(this SheetCustomization sheetInfo, Color color)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);
        InitializeFontStyle(sheetInfo);

        sheetInfo.Style.SetFontColor(color);
        return sheetInfo;
    }

    public static SheetCustomization SetFontHeightInPoints(this SheetCustomization sheetInfo, short height)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);
        InitializeFontStyle(sheetInfo);

        sheetInfo.Style.SetFontHeightInPoints(height);
        return sheetInfo;
    }

    public static SheetCustomization SetFontBold(this SheetCustomization sheetInfo, bool isBold)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);
        InitializeFontStyle(sheetInfo);

        sheetInfo.Style.SetFontBold(isBold);
        return sheetInfo;
    }
    public static SheetCustomization SetDateTimeFormat(this SheetCustomization sheetInfo, string format)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        InitializeStyle(sheetInfo);

        sheetInfo.Style.SetDateTimeFormat(format);
        return sheetInfo;
    }


    private static void InitializeStyle(SheetCustomization sheetInfo)
    {
        if (sheetInfo.Style is null)
        {
            sheetInfo.Style = new();
        }
    }

    private static void InitializeFontStyle(SheetCustomization sheetInfo)
    {
        if (sheetInfo.Style.FontStyle is null)
        {
            sheetInfo.Style.FontStyle = new();
        }
    }
}
