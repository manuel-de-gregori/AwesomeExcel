using AwesomeExcel.Common.Models;

namespace AwesomeExcel.Common.Services;

/// <summary>
/// Merge common Styles into one single Style.
/// </summary>
public class StylesMerger
{
    /// <summary>
    /// Merge all the given styles into one new Style.
    /// </summary>
    /// <param name="styles">The styles to be merged. 
    /// <br /> Lower priority to the left elements. </param>
    /// <returns>A new style with all the informations of the given styles.</returns>
    public Style Merge(params Style[] styles)
    {
        Style accumulator = null;
        foreach (Style current in styles)
        {
            if (current == null)
                continue;
            else if (accumulator == null)
                accumulator = DeepCopy(current);
            else
                Merge(accumulator, current);
        }

        return accumulator;
    }

    private Style DeepCopy(Style style)
    {
        Style result = style.ShallowCopy();
        result.FontStyle = style.FontStyle?.ShallowCopy();
        return result;
    }

    private void Merge(Style accumulator, Style current)
    {
        SetBorderTopColor();
        SetBorderBottomColor();
        SetBorderLeftColor();
        SetBorderRightColor();
        SetFillPattern();
        SetFillForegroundColor();
        SetFontStyle();
        SetHorizontalAlignment();
        SetVerticalAlignment();
        SetDateTimeFormat();

        void SetBorderTopColor()
        {
            if (current.BorderTopColor.HasValue)
                accumulator.BorderTopColor = current.BorderTopColor.Value;
        }

        void SetBorderBottomColor()
        {
            if (current.BorderBottomColor.HasValue)
                accumulator.BorderBottomColor = current.BorderBottomColor.Value;
        }
        void SetBorderLeftColor()
        {
            if (current.BorderLeftColor.HasValue)
                accumulator.BorderLeftColor = current.BorderLeftColor.Value;
        }

        void SetBorderRightColor()
        {
            if (current.BorderRightColor.HasValue)
                accumulator.BorderRightColor = current.BorderRightColor.Value;
        }

        void SetFillPattern()
        {
            if (current.FillPattern.HasValue)
                accumulator.FillPattern = current.FillPattern.Value;
        }

        void SetFillForegroundColor()
        {
            if (current.FillForegroundColor.HasValue)
                accumulator.FillForegroundColor = current.FillForegroundColor.Value;
        }

        void SetFontStyle()
        {
            if (current.FontStyle != null)
            {
                if (accumulator.FontStyle == null)
                {
                    accumulator.FontStyle = current.FontStyle.ShallowCopy();
                }
                else
                {
                    Merge(accumulator.FontStyle, current.FontStyle);
                }
            }
        }

        void SetHorizontalAlignment()
        {
            if (current.HorizontalAlignment.HasValue)
                accumulator.HorizontalAlignment = current.HorizontalAlignment.Value;
        }

        void SetVerticalAlignment()
        {
            if (current.VerticalAlignment.HasValue)
                accumulator.VerticalAlignment = current.VerticalAlignment.Value;
        }

        void SetDateTimeFormat()
        {
            if (!string.IsNullOrWhiteSpace(current.DateTimeFormat))
                accumulator.DateTimeFormat = current.DateTimeFormat;
        }
    }

    private void Merge(FontStyle accumulator, FontStyle current)
    {
        SetName();
        SetColor();
        SetHeightInPoints();
        SetIsBold();

        void SetName()
        {
            if (!string.IsNullOrWhiteSpace(current.Name)) 
                accumulator.Name = current.Name;
        }

        void SetColor()
        {
            if (current.Color.HasValue) 
                accumulator.Color = current.Color.Value;
        }

        void SetHeightInPoints()
        {
            if (current.HeightInPoints.HasValue) 
                accumulator.HeightInPoints = current.HeightInPoints.Value;
        }

        void SetIsBold()
        {
            if (current.IsBold.HasValue) 
                accumulator.IsBold = current.IsBold.Value;
        }
    }
}
