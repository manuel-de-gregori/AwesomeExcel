using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using AwesomeExcel.Customization.Services;
using AwesomeExcel.FluentCustomization;
using AwesomeExcel.Generator.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace AwesomeExcel.Generator.UnitTests;

[TestClass]
public class SheetFactory_CellsCustomizationTest
{
    private readonly List<Person> data = new()
    {
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
    };

    [TestMethod]
    public void CreateSheet_NullChecks()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        Assert.IsNotNull(sheet);
        Assert.IsNotNull(sheet.Columns);
        Assert.IsNotNull(sheet.Rows);
        Assert.IsNull(sheet.Style);
        Assert.IsNull(sheet.HeaderStyle);

        Cell cell0 = sheet.Rows[0].Cells[0];
        Assert.IsNotNull(cell0.Style);

        Column column1 = sheet.Columns[1];
        Assert.IsNotNull(column1.Style);

        Column column2 = sheet.Columns[2];
        Assert.IsNotNull(column2.Style?.FontStyle);

        Column column3 = sheet.Columns[3];
        Assert.IsNotNull(column3.Style?.FontStyle);
    }

    [TestMethod]
    public void CreateSheet_ColumnName_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        foreach (Row row in sheet.Rows)
        {
            const int columnName = 0;
            Cell cell = row.Cells[columnName];

            Assert.AreEqual(cell.Style.HorizontalAlignment, HorizontalAlignment.Right);
        }
    }

    [TestMethod]
    public void CreateSheet_ColumnSurname_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        foreach (Row row in sheet.Rows)
        {
            const int columnSurname = 1;
            Cell cell = row.Cells[columnSurname];
            string surname = (string)cell.Value;

            Color? expected = surname == "da Vinci"
                ? Color.Blue
                : surname == "DiCaprio"
                    ? Color.Green
                    : null;

            Color? actual = cell.Style.FillForegroundColor;

            Assert.AreEqual(actual, expected);
        }
    }

    [TestMethod]
    public void CreateSheet_ColumnBirthDate_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        Column column2 = sheet.Columns[2];
        Assert.IsTrue(column2.Style.FontStyle.IsBold);
        Assert.AreEqual(column2.Style.FontStyle.Color, Color.Red);
    }

    [TestMethod]
    public void CreateSheet_ColumnAge_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        Column column3 = sheet.Columns[3];
        Assert.AreEqual(actual: column3.Style.BorderBottomColor, expected: Color.Green);
        Assert.AreEqual(actual: column3.Style.BorderLeftColor, expected: Color.IceBlue);
        Assert.AreEqual(actual: column3.Style.BorderRightColor, expected: Color.Lime);
        Assert.AreEqual(actual: column3.Style.BorderTopColor, expected: Color.Red);
        Assert.AreEqual(actual: column3.Style.FillForegroundColor, expected: Color.Ivory);
        //Assert.AreEqual(actual: column3.Style.FillPattern, expected: FillPattern.SolidForeground);
        Assert.AreEqual(actual: column3.Style.HorizontalAlignment, expected: HorizontalAlignment.Right);
        Assert.AreEqual(actual: column3.Style.VerticalAlignment, expected: VerticalAlignment.Bottom);
        //Assert.AreEqual(actual: fourhColumn.Style.DateTimeFormat, expected: "a5b2");

        Assert.AreEqual(actual: column3.Style.FontStyle.Color, expected: Color.Yellow);
        Assert.AreEqual(actual: column3.Style.FontStyle.HeightInPoints, expected: (short)12);
        Assert.AreEqual(actual: column3.Style.FontStyle.IsBold, expected: null);
        Assert.AreEqual(actual: column3.Style.FontStyle.Name, expected: "AwesomeExcel");
    }

    private class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age => (DateTime.Now.Date - BirthDate.Date).Days / 365;
    }

    public Dictionary<PropertyInfo, CellCustomization> GetCustomizedCells()
    {
        return new Dictionary<PropertyInfo, CellCustomization>
        {
            { typeof(Person).GetProperty(nameof(Person.Name)), GetColumnName() },
            { typeof(Person).GetProperty(nameof(Person.Surname)), GetColumnSurname() },
            { typeof(Person).GetProperty(nameof(Person.BirthDate)), GetColumnBirthDate() },
            { typeof(Person).GetProperty(nameof(Person.Age)), GetColumnAge() },
        };
    }

    private CellCustomization GetColumnAge()
    {
        CellCustomization<int> customizedColumn = new();

        customizedColumn
            .SetBorderLeftColor(value => Color.IceBlue)
            .SetBorderRightColor(value => Color.Lime)
            .SetBorderTopColor(value => Color.Red)
            .SetBorderBottomColor(value => Color.Green)
            .SetFontHeightInPoints(value => 12)
            .SetFontColor(value => Color.Yellow)
            .SetHorizontalAlignment(value => HorizontalAlignment.Right)
            .SetVerticalAlignment(value => VerticalAlignment.Bottom)
            .SetFillForegroundColor(value => Color.Ivory)
            .SetFontName(value => "AwesomeExcel");


        return customizedColumn;
    }

    private CellCustomization GetColumnBirthDate()
    {
        CellCustomization<DateTime> customizedColumn = new();

        customizedColumn
            .SetFontBold(value => true)
            .SetFontColor(value => Color.Red);

        return customizedColumn;
    }

    private CellCustomization GetColumnSurname()
    {
        CellCustomization<string> customizedColumn = new();
        customizedColumn.SetFillForegroundColor(value =>
        {
            if (value == "da Vinci")
                return Color.Blue;

            if (value == "DiCaprio")
                return Color.Green;

            return null;
        });
        return customizedColumn;
    }

    private CellCustomization<string> GetColumnName()
    {
        CellCustomization<string> _customizedColumn = new();

        _customizedColumn.SetHorizontalAlignment(s => HorizontalAlignment.Right);

        return _customizedColumn;
    }
}
