using AwesomeExcel.Common.Models;

namespace AwesomeExcel.Customization.Models;

public class StyleCustomization { }

public class StyleCustomization<T> : StyleCustomization
{
    public Func<T, Color?> BorderTopColor { get; set; }
    public Func<T, Color?> BorderBottomColor { get; set; }
    public Func<T, Color?> BorderLeftColor { get; set; }
    public Func<T, Color?> BorderRightColor { get; set; }
    public Func<T, Color?> FillForegroundColor { get; set; }
    public Func<T, FillPattern?> FillPattern { get; set; }
    public Func<T, string> DateTimeFormat { get; set; }
    public Func<T, HorizontalAlignment?> HorizontalAlignment { get; set; }
    public Func<T, VerticalAlignment?> VerticalAlignment { get; set; }

    public FontStyleCustomization<T> FontStyle { get; set; }
}
