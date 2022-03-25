using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Services;
using AwesomeExcel.FluentCustomization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace AwesomeExcel.IntegrationTests;

[TestClass]
public class Class2
{
    [TestMethod]
    public void TestMethod1()
    {
        List<Person> people = GetActors();

        var fileGenerator = new BridgeNpoi.NpoiFileGenerator();

        MemoryStream file = fileGenerator.Generate(people, (ISheetCustomizer<Person> sps) =>
        {
            sps.Workbook.SetFileType(FileType.Xlsx);

            sps.Sheet
                .SetName("Test nome foglio")
                .SetFillForegroundColor(Color.LightBlue)
                .SetHeaderFillForegroundColor(Color.Blue)
                .SetHeaderBorderBottomColor(Color.Red)
                .SetVerticalAlignment(VerticalAlignment.Center);

            sps.Column(p => p.Name)
                .SetName("Actor's name")
                .SetStyle(s => s.FillForegroundColor = Color.Aqua);

            sps.Column(p => p.Surname)
                .SetName("Actor's surname")
                .SetHorizontalAlignment(HorizontalAlignment.Right);

            sps.Column(p => p.BirthDate)
                .SetStyle(s => s.DateTimeFormat = "dd/mm/yyyy");

            sps.Cells(p => p.BirthDate)
                .SetFillForegroundColor(birthDate => birthDate.HasValue && birthDate.Value.Month == 3 ? Color.Red : Color.Yellow);
        });

        string fileName = nameof(Class2) + "-" + nameof(TestMethod1) + ".xlsx";
        WriteFile(file, fileName);
    }

    [TestMethod]
    public void TestMethod2()
    {
        List<Person> people = GetActors();
        List<Invoice> invoices = GetInvoices();

        var fileGenerator = new BridgeNpoi.NpoiFileGenerator();

        MemoryStream file = fileGenerator.Generate(people, invoices, (ISheetsCustomizer<Person, Invoice> sps) =>
        { 
            sps.Workbook.SetFileType(FileType.Xlsx);

            sps.Sheet1
                .SetName("Test sheet name")
                .SetFillForegroundColor(Color.LightBlue)
                .SetHeaderFillForegroundColor(Color.Blue)
                .SetHeaderBorderBottomColor(Color.Red)
                .SetVerticalAlignment(VerticalAlignment.Center);

            sps.Sheet2.HasHeader(true);

            sps.Column(sps.Sheet1, p => p.Name)
                .SetName("Actor's name")
                .SetStyle(s => s.FillForegroundColor = Color.Aqua);

            sps.Column(sps.Sheet1, p => p.Surname)
                .SetName("Actor's surname")
                .SetHorizontalAlignment(HorizontalAlignment.Right);

            sps.Column(sps.Sheet2, p => p.CreationDate)
                .SetDateTimeFormat("dd/mm/yyyy");

            sps.Column(sps.Sheet2, p => p.Amount)
                .SetFillForegroundColor(Color.Green);
        });

        string fileName = nameof(Class2) + "-" + nameof(TestMethod2) + ".xlsx";
        WriteFile(file, fileName);
    }

    [TestMethod]
    public void TestMethod3()
    {
        List<Person> people = GetActors();
        List<Invoice> invoices = GetInvoices();

        var fileGenerator = new BridgeNpoi.NpoiFileGenerator();

        MemoryStream file = fileGenerator.Generate(people, invoices, (c) => { });

        string fileName = nameof(Class2) + "-" + nameof(TestMethod2) + ".xlsx";
        WriteFile(file, fileName);
    }


    //private List<Person> GetActorsFromFile()
    //{
    //    string path = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "resources", "actors.json");
    //    string content = System.IO.File.ReadAllText(path);
    //    List<Person> people = JsonSerializer.Deserialize<List<Person>>(content, new JsonSerializerOptions(JsonSerializerDefaults.Web) { });
    //    return people;
    //}

    private List<Person> GetActors()
    {
        return new List<Person>
        {
            { new() { Name =  "Caroline", Surname = "Aaron", BirthDate = DateTime.Parse("1952-08-07") } },
            { new() { Name =  "Victor", Surname = "Aaron", BirthDate = DateTime.Parse("1956-09-11") } },
            { new() { Name =  "Diego", Surname = "Abatantuono", BirthDate = DateTime.Parse("1955-05-20") } },
            { new() { Name =  "Andrew", Surname = "Abeita", BirthDate = DateTime.Parse("1981-07-11") } },
            { new() { Name =  "Jon", Surname = "Abrahams", BirthDate = DateTime.Parse("1977-10-29") } },
            { new() { Name =  "Stefano", Surname = "Accorsi", BirthDate = DateTime.Parse("1971-03-02") } },
            { new() { Name =  "Dean", Surname = "Acheson", BirthDate = DateTime.Parse("1893-04-11") } },
            { new() { Name =  "Josh", Surname = "Ackerman", BirthDate = DateTime.Parse("1977-03-23") } },
            { new() { Name =  "Joss", Surname = "Ackland", BirthDate = DateTime.Parse("1928-02-29") } },
            { new() { Name =  "Jay", Surname = "Acovone", BirthDate = DateTime.Parse("1955-08-20") } },
            { new() { Name =  "Deb", Surname = "Adair", BirthDate = DateTime.Parse("1966-04-22") } },
            { new() { Name =  "Enid-Raye", Surname = "Adams", BirthDate = DateTime.Parse("1973-06-16") } },
            { new() { Name =  "Jacob", Surname = "Adams", BirthDate = DateTime.Parse("1975-07-04") } },
            { new() { Name =  "Mario", Surname = "Adorf", BirthDate = DateTime.Parse("1930-09-08") } },
            { new() { Name =  "Ben", Surname = "Affleck", BirthDate = DateTime.Parse("1972-08-15") } },
            { new() { Name =  "Casey", Surname = "Affleck", BirthDate = DateTime.Parse("1975-08-12") } },
            { new() { Name =  "Spiro", Surname = "Agnew", BirthDate = DateTime.Parse("1918-11-09") } },
            { new() { Name =  "Antonio", Surname = "Agri", BirthDate = DateTime.Parse("1932-05-05") } },
            { new() { Name =  "Jenny", Surname = "Agutter", BirthDate = DateTime.Parse("1952-12-20") } },
            { new() { Name =  "Betsy", Surname = "Aidem", BirthDate = DateTime.Parse("1957-10-28") } },
            { new() { Name =  "Liam", Surname = "Aiken", BirthDate = DateTime.Parse("1990-01-07") } },
            { new() { Name =  "Troy", Surname = "Aikman", BirthDate = DateTime.Parse("1966-11-21") } },
            { new() { Name =  "Kacey", Surname = "Ainsworth", BirthDate = DateTime.Parse("1970-10-19") } },
            { new() { Name =  "Holly", Surname = "Aird", BirthDate = DateTime.Parse("1969-05-18") } },
            { new() { Name =  "Lucy", Surname = "Akhurst", BirthDate = DateTime.Parse("1975-11-18") } },
            { new() { Name =  "Amy", Surname = "Alcott", BirthDate = DateTime.Parse("1956-02-22") } },
            { new() { Name =  "Alan", Surname = "Alda", BirthDate = DateTime.Parse("1936-01-28") } },
            { new() { Name =  "Tom", Surname = "Aldredge", BirthDate = DateTime.Parse("1928-02-28") } },
            { new() { Name =  "Buzz", Surname = "Aldrin", BirthDate = DateTime.Parse("1930-01-20") } },
            { new() { Name =  "Henry", Surname = "Alessandroni", BirthDate = DateTime.Parse("1959-05-26") } },
            { new() { Name =  "Art", Surname = "Alexakis", BirthDate = DateTime.Parse("1962-04-12") } },
            { new() { Name =  "Jane", Surname = "Alexander", BirthDate = DateTime.Parse("1939-10-28") } },
            { new() { Name =  "Jason", Surname = "Alexander", BirthDate = DateTime.Parse("1959-09-23") } },
            { new() { Name =  "Khandi", Surname = "Alexander", BirthDate = DateTime.Parse("1957-09-04") } },
            { new() { Name =  "Adam", Surname = "Alexi-Malle", BirthDate = DateTime.Parse("1964-09-24") } },
            { new() { Name =  "Hans", Surname = "Alfredson", BirthDate = DateTime.Parse("1931-06-28") } },
            { new() { Name =  "Mary", Surname = "Alice", BirthDate = DateTime.Parse("1941-12-03") } },
            { new() { Name =  "Debbie", Surname = "Allen", BirthDate = DateTime.Parse("1950-01-16") } }
        };
    }

    private List<Invoice> GetInvoices()
    {
        return new List<Invoice>
        {
            { new() { Amount = 35, CreationDate = new DateTime(2016, 01, 01) } },
            { new() { Amount = 52, CreationDate = new DateTime(2016, 02, 01) } },
            { new() { Amount = 123123, CreationDate = new DateTime(2016, 03, 01) } },
            { new() { Amount = 34345654, CreationDate = new DateTime(2016, 04, 01) } },
            { new() { Amount = 234, CreationDate = new DateTime(2016, 05, 01) } },
            { new() { Amount = 123, CreationDate = new DateTime(2016, 06, 01) } },
            { new() { Amount = 35, CreationDate = new DateTime(2016, 06, 05) } },
            { new() { Amount = 7, CreationDate = new DateTime(2016, 06, 12) } },
            { new() { Amount = 3567, CreationDate = new DateTime(2018, 01, 01) } },
            { new() { Amount = 8776, CreationDate = new DateTime(2019, 01, 01) } },
            { new() { Amount = 567, CreationDate = new DateTime(2020, 01, 01) } },
        };
    }

    private void WriteFile(MemoryStream file, string fileName)
    {
        string directory = @"C:\\Users\\manue\\OneDrive\\Documents\\Repositories\\test-excel";
        string path = Path.Combine(directory, fileName);
        byte[] fileBytes = file.ToArray();
        File.WriteAllBytes(path, fileBytes);
    }
}

internal class Person
{
    public string Name { get; set; }
    public string Surname { get; set; }
    public DateTime? BirthDate { get; set; }
}

internal class Invoice
{
    public DateTime CreationDate { get; set; }
    public double Amount { get; set; }
}