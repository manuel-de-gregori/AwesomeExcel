using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AwesomeExcel.Customization.Services;

public class CellStyleCustomizationResolver
{
    public Style Resolve(CellCustomization customization, object value)
    {
        CellStyleCustomization csc = (CellStyleCustomization)customization.GetType().GetProperty("Style").GetValue(customization);

        if (csc is null)
            return null;

        Color? borderTopColor = GetBorderTopColor(csc, value);
        Color? borderBottomColor = GetBorderBottomColor(csc, value);
        Color? borderLeftColor = GetBorderLeftColor(csc, value);
        Color? borderRightColor = GetBorderRightColor(csc, value);
        Color? fillForegroundColor = GetFillForegroundColor(csc, value);
        FillPattern? fillPattern = GetFillPattern(csc, value);
        string dateTimeFormat = GetDateTimeFormat(csc, value);
        HorizontalAlignment? horizontalAlignment = GetHorizontalAlignment(csc, value);
        VerticalAlignment? verticalAlignment = GetVerticalAlignment(csc, value);
        FontStyle fontStyle = GetFontStyle(value, csc);

        return new Style
        {
            BorderTopColor = borderTopColor,
            BorderBottomColor = borderBottomColor,
            BorderLeftColor = borderLeftColor,
            BorderRightColor = borderRightColor,
            FillForegroundColor = fillForegroundColor,
            FillPattern = fillPattern,
            DateTimeFormat = dateTimeFormat,
            HorizontalAlignment = horizontalAlignment,
            VerticalAlignment = verticalAlignment,
            FontStyle = fontStyle
        };
    }

    private FontStyle GetFontStyle(object value, CellStyleCustomization sc)
    {
        CellFontStyleCustomization cfsc = GetFontStyleCustomization(sc, value);

        if (cfsc is null)
            return null;

        Color? color = GetFontColor(cfsc, value);
        short? heightInPoints = GetFontHeightInPoints(cfsc, value);
        bool? isBold = GetFontBold(cfsc, value);
        string name = GetFontName(cfsc, value);

        return new FontStyle
        {
            Color = color,
            HeightInPoints = heightInPoints,
            IsBold = isBold,
            Name = name
        };
    }

    private Color? GetBorderTopColor(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderTopColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetBorderBottomColor(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderBottomColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetBorderLeftColor(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderLeftColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetBorderRightColor(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderRightColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetFillForegroundColor(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.FillForegroundColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private FillPattern? GetFillPattern(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.FillPattern);
        return GetValue<FillPattern?>(sc, pName, value);
    }

    private HorizontalAlignment? GetHorizontalAlignment(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.HorizontalAlignment);
        return GetValue<HorizontalAlignment?>(sc, pName, value);
    }

    private VerticalAlignment? GetVerticalAlignment(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.VerticalAlignment);
        return GetValue<VerticalAlignment?>(sc, pName, value);
    }

    private string GetDateTimeFormat(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.DateTimeFormat);
        return GetValue<string>(sc, pName, value);
    }

    private CellFontStyleCustomization GetFontStyleCustomization(CellStyleCustomization sc, object value)
    {
        const string pName = nameof(CellStyleCustomization<object>.FontStyle);
        return GetValue<CellFontStyleCustomization>(sc, pName, value);
    }

    private Color? GetFontColor(CellFontStyleCustomization cfsc, object value)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.Color);
        return GetFontValue<Color?>(cfsc, pName, value);
    }

    private short? GetFontHeightInPoints(CellFontStyleCustomization cfsc, object value)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.HeightInPoints);
        return GetFontValue<short?>(cfsc, pName, value);
    }

    private bool? GetFontBold(CellFontStyleCustomization cfsc, object value)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.IsBold);
        return GetFontValue<bool?>(cfsc, pName, value);
    }

    private string GetFontName(CellFontStyleCustomization cfsc, object value)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.Name);
        return GetFontValue<string>(cfsc, pName, value);
    }

    private T GetFontValue<T>(CellFontStyleCustomization cfsc, string pName, object value)
    {
        Type fscType = typeof(CellFontStyleCustomization<object>);
        PropertyInfo pi = fscType.GetProperty(pName);
        var pValue = (Delegate)pi.GetValue(cfsc);
        var result = Invoke<T>(pValue, value);
        return result;
    }

    private T GetValue<T>(CellStyleCustomization sc, string pName, object value)
    {
        Type type = typeof(CellStyleCustomization<object>);
        PropertyInfo pi = type.GetProperty(pName);
        var pValue = (Delegate)pi.GetValue(sc);
        var result = Invoke<T>(pValue, value);
        return result;
    }

    private T Invoke<T>(Delegate fn, object value)
    {
        return (T)fn?.Method.Invoke(fn.Target, new[] { value });
    }
}
