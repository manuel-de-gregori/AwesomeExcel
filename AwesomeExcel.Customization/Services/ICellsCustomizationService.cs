using AwesomeExcel.Customization.Models;
using System.Reflection;

namespace AwesomeExcel.Customization.Services
{
    public interface ICellsCustomizationService
    {
        CellCustomization GetCells(PropertyInfo pi);
    }
}