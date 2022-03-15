using AwesomeExcel.Customization.Models;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

/* to do: rename these types */

public interface IColumnCustomizationService
{
    ColumnCustomization GetColumn(PropertyInfo pi);
}
