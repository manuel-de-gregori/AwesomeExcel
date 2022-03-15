﻿using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public abstract class MultipleSheetsCustomizer
{
    public WorkbookCustomization Workbook { get; } = new();

    internal Dictionary<SheetCustomization, ColumnsCustomizer> dict = new();

    public ColumnCustomization GetOrCreateColumn<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        if (!dict.ContainsKey(sheet))
        {
            ColumnsCustomizer ccs = new();
            dict.Add(sheet, ccs);
        }

        var ccs2 = dict[sheet];

        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        return ccs2.GetOrCreateColumn(pi);
    }

    public Cell<TProperty> GetCells<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        return null;
    }

    public ColumnsCustomizer GetColumnCustomizationService<T>(SheetCustomization<T> sheet)
    {
        ColumnsCustomizer ccs = dict[sheet];
        return ccs;
    }
}

public class MultipleSheetsCustomizationService<TSheet1, TSheet2> : MultipleSheetsCustomizer
{
    public (SheetCustomization<TSheet1> Sheet1, SheetCustomization<TSheet2> Sheet2) Sheets { get; } = new(new(), new());
}

public class MultipleSheetsCustomizationService<TSheet1, TSheet2, TSheet3> : MultipleSheetsCustomizer
{
    public (SheetCustomization<TSheet1>, SheetCustomization<TSheet2>, SheetCustomization<TSheet3>) Sheets { get; } = new(new(), new(), new());
}

public class MultipleSheetsCustomizationService<TSheet1, TSheet2, TSheet3, TSheet4> : MultipleSheetsCustomizer
{
    public (SheetCustomization<TSheet1>, SheetCustomization<TSheet2>, SheetCustomization<TSheet3>, SheetCustomization<TSheet4>) Sheets { get; } = new(new(), new(), new(), new());
}

public class MultipleSheetsCustomizationService<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5> : MultipleSheetsCustomizer
{
    public (SheetCustomization<TSheet1>, SheetCustomization<TSheet2>, SheetCustomization<TSheet3>, SheetCustomization<TSheet4>, SheetCustomization<TSheet5>) Sheets { get; } = new(new(), new(), new(), new(), new());
}