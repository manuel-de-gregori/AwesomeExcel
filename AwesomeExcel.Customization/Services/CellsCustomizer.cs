using AwesomeExcel.Customization.Models;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public class CellsCustomizer<T>
{
    private readonly Dictionary<PropertyInfo, CellCustomization> customizedCells = new();

    public CellCustomization GetCells(PropertyInfo pi)
    {
        customizedCells.TryGetValue(pi, out CellCustomization value);
        return value;
    }

    public CellCustomization GetOrCreateCells(PropertyInfo pi)
    {
        if (customizedCells.TryGetValue(pi, out CellCustomization value))
        {
            return value;
        }
        else
        {
            CellCustomization ci = new();
            customizedCells.Add(pi, ci);
            return ci;
        }
    }
}
