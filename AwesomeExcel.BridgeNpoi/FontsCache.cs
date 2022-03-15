using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class FontsCache
{
    private readonly _NPOI.IFont emptyFont;
    private readonly Dictionary<_Excel.FontStyle, _NPOI.IFont> cache = new(new Common.Comparers.FontStyleEqualityComparer());
    private readonly Dictionary<_Excel.FontStyle, int> referenceCounter = new(new Common.Comparers.FontStyleEqualityComparer());

    public FontsCache(_NPOI.IWorkbook npoiWorkbook)
    {
        emptyFont = npoiWorkbook.CreateFont();
    }

    public _NPOI.IFont Get(_Excel.FontStyle fontStyle)
    {
        // Problem 1:
        //     This cache is necessary because there's a limit on how many Fonts can be used in a workbook.
        //
        //     Explanation:
        //         "the maximum number of unique fonts in a workbook is limited to 32767. You should re-use fonts in your applications instead of creating a font for each cell"
        //
        //     Official source:
        //         https://poi.apache.org/components/spreadsheet/quick-guide.html#WorkingWithFonts
        //
        //     Solution:
        //        Using a cache to re-use the same instance of (NPOI) IFont for multiple cells/rows
        // 
        //
        // Problem 2:
        //    It's necessary to keep count of how many references of the same FontStyle have been used.
        //    Because there's a limit on how many times one (NPOI) IFont instance can be used in a workbook.
        //
        //    Explanation:
        //        the maximum number of times that an IFont instance can be used in a workbook is limited to 24.
        //
        //    Official source:
        //        none
        //
        //    Solution:    
        //        This cache keep count of how many references of the same font have been used.


        if (fontStyle == null)
            return emptyFont;

        (_NPOI.IFont npoiFont, int referenceCounter) = GetFromCache(fontStyle);

        if (npoiFont != null)
        {
            // One IFont instance can be used for styling up to 24 times 
            bool limitReached = referenceCounter >= 25;

            if (limitReached)
            {
                // Remove the font from cache.
                // In this way, a new font instance will be created for this font style
                Remove(fontStyle);
                return null;
            }
            else
            {
                return npoiFont;
            }
        }

        return null;
    }

    private (_NPOI.IFont instance, int usageCount) GetFromCache(_Excel.FontStyle fontStyle)
    {
        if (cache.TryGetValue(fontStyle, out _NPOI.IFont npoiFont))
        {
            referenceCounter[fontStyle]++;
            return (npoiFont, referenceCounter[fontStyle]);
        }

        return (null, 0);
    }

    public void Add(_NPOI.IFont npoiFont, _Excel.FontStyle fontStyle)
    {
        cache.Add(fontStyle, npoiFont);
        referenceCounter.Add(fontStyle, 0);
    }

    private void Remove(_Excel.FontStyle fontStyle)
    {
        cache.Remove(fontStyle);
        referenceCounter.Remove(fontStyle);
    }

    public void Clear()
    {
        cache.Clear();
        referenceCounter.Clear();
    }
}