# AwesomeExcel

```csharp
fileGenerator.Generate(people, (customizer) =>
{
    customizer.Workbook.SetFileType(FileType.Xlsx);

    customizer.Sheet
        .SetName("Test nome foglio")
        .SetFillForegroundColor(Color.LightBlue)
        .SetHeaderFillForegroundColor(Color.Blue)
        .SetHeaderBorderBottomColor(Color.Red)
        .SetVerticalAlignment(VerticalAlignment.Center);

    customizer.GetColumn(p => p.Name)
        .SetName("Actor's name")
        .SetStyle(s => s.FillForegroundColor = Color.Aqua);

    customizer.GetColumn(p => p.Surname)
        .SetName("Actor's surname")
        .SetHorizontalAlignment(HorizontalAlignment.Right);

    customizer.GetColumn(p => p.BirthDate)
        .SetStyle(s => s.DateTimeFormat = "dd/mm/yyyy");
});
```
