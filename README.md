# AwesomeExcel

```csharp
var stream = excelGenerator.Generate(actors, (customizer) =>
{
    customizer.Sheet
        .SetName("Sheet name for README example")
        .SetFillForegroundColor(Color.LightBlue)
        .SetVerticalAlignment(VerticalAlignment.Center);

    customizer.Column(p => p.Name)
        .SetName("Actor's name")
        .SetFillForegroundColor(Color.Aqua);

    customizer.Column(p => p.Surname)
        .SetName("Actor's surname")
        .SetHorizontalAlignment(HorizontalAlignment.Right);

    customizer.Column(p => p.BirthDate)
        .SetDateTimeFormat("dd/mm/yyyy");
});

class Actor
{
    public string Name { get; set; }
    public string Surname { get; set; }
    public DateTime? BirthDate { get; set; }
}
```
