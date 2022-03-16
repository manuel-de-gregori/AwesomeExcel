# AwesomeExcel

```csharp
var stream = excelGenerator.Generate(actors, (customizer) =>
{
    customizer.Workbook.SetFileType(FileType.Xlsx);

    customizer.Sheet
        .SetName("Sheet name for README example")
        .SetFillForegroundColor(Color.LightBlue)
        .SetVerticalAlignment(VerticalAlignment.Center);

    customizer.GetColumn(p => p.Name)
        .SetName("Actor's name")
        .SetFillForegroundColor(Color.Aqua);

    customizer.GetColumn(p => p.Surname)
        .SetName("Actor's surname")
        .SetHorizontalAlignment(HorizontalAlignment.Right);

    customizer.GetColumn(p => p.BirthDate)
        .SetDateTimeFormat("dd/mm/yyyy");
});

class Actor
{
    public string Name { get; set; }
    public string Surname { get; set; }
    public DateTime? BirthDate { get; set; }
}
```
