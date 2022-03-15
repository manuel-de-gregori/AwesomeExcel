using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.FluentCustomization;

public static class SheetCustomizationExtension
{
    public static SheetCustomization SetName(this SheetCustomization sheetInfo, string name)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        sheetInfo.Name = name;
        return sheetInfo;
    }

    public static SheetCustomization Protect(this SheetCustomization sheetInfo)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        sheetInfo.IsProtected = true;
        return sheetInfo;
    }

    public static SheetCustomization HasHeader(this SheetCustomization sheetInfo, bool hasHeader)
    {
        if (sheetInfo is null)
        {
            throw new ArgumentNullException(nameof(sheetInfo));
        }

        sheetInfo.HasHeader = hasHeader;
        return sheetInfo;
    }
}
