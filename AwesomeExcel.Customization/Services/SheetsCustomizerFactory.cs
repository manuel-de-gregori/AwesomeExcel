namespace AwesomeExcel.Customization.Services;

internal class SheetsCustomizerFactory
{
    public ISheetCustomizer<TSheet> Get<TSheet>()
    {
        return new SheetCustomizer<TSheet>();
    }

    public ISheetsCustomizer<TSheet1, TSheet2> Get<TSheet1, TSheet2>()
    {
        return new SheetsCustomizer<TSheet1, TSheet2>();
    }

    public ISheetsCustomizer<TSheet1, TSheet2, TSheet3> Get<TSheet1, TSheet2, TSheet3>()
    {
        return new MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3>();
    }

    public ISheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4> Get<TSheet1, TSheet2, TSheet3, TSheet4>()
    {
        return new MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4>();
    }

    public ISheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5> Get<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5>()
    {
        return new MultipleSheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5>();
    }
}
