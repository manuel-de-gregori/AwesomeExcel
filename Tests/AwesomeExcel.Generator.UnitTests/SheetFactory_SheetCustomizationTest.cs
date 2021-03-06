using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using AwesomeExcel.Generator.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AwesomeExcel.Generator.UnitTests;

[TestClass]
public class SheetFactory_SheetCustomizationTest
{
    private readonly List<Person> data = new()
    {
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) }
    };

    [TestMethod]
    public void CreateSheet_ShouldHaveGivenValues()
    {
        SheetCustomization<Person> si = new()
        {
            Name = nameof(CreateSheet_ShouldHaveGivenValues),
            HasHeader = true,
            IsProtected = true
        };

        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, si, null, null);

        Assert.AreEqual(sheet.Name, nameof(CreateSheet_ShouldHaveGivenValues));
        Assert.AreEqual(sheet.HasHeader, true);
        Assert.AreEqual(sheet.IsProtected, true);
    }

    [TestMethod]
    public void CreateSheet_ShouldHaveGivenValues_2()
    {
        SheetCustomization<Person> si = new()
        {
            Name = nameof(CreateSheet_ShouldHaveGivenValues_2),
            HasHeader = false,
            IsProtected = false
        };


        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, si, null, null);

        Assert.AreEqual(sheet.Name, nameof(CreateSheet_ShouldHaveGivenValues_2));
        Assert.AreEqual(sheet.HasHeader, false);
        Assert.AreEqual(sheet.IsProtected, false);
    }

    [TestMethod]
    public void CreateSheet_StyleAndFontStyle_ShouldHaveGivenValues()
    {
        SheetCustomization<Person> si = new()
        {
            Style = new()
            {
                BorderBottomColor = Color.Aqua,
                BorderTopColor = Color.Red
            },
            HeaderStyle = new()
            {
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom
            }
        };

        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, si, null, null);

        Assert.AreEqual(sheet.Style.BorderBottomColor, Color.Aqua);
        Assert.AreEqual(sheet.Style.BorderTopColor, Color.Red);

        Assert.AreEqual(sheet.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Right);
        Assert.AreEqual(sheet.HeaderStyle.VerticalAlignment, VerticalAlignment.Bottom);
    }

    private class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age => (DateTime.Now.Date - BirthDate.Date).Days / 365;
    }
}
