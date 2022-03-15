namespace AwesomeExcel.Customization.Models;

public class CellCustomization
{

}

public class CellCustomization<TProperty> : CellCustomization
{
    public StyleCustomization<TProperty> Style { get; set; }
}