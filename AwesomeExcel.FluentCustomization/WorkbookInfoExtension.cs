using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.FluentCustomization;

public static class WorkbookInfoExtension
{
    public static WorkbookCustomization SetFileType(this WorkbookCustomization workbook, FileType fileType)
    {
        workbook.FileType = fileType;
        return workbook;
    }
}
