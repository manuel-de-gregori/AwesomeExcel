using AwesomeExcel.BridgeNpoi;
using AwesomeExcel.Common.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AwesomeExcel.IntegrationTests;

[TestClass]
public class Class1
{
    [TestMethod]
    public void TestMethod1()
    {
        Workbook excelWorkbook = new()
        {
            FileType = FileType.Xlsx,
            Sheets = new Sheet[1]
            {
                new()
                {
                    Name = "Daniel LaRusso",
                    Rows = new Row[1]
                    {
                        new()
                        {
                            Style = new()
                            {
                                FontStyle = null
                            },
                            Cells = new Cell[2]
                            {
                                new()
                                {
                                    Value = null,
                                    Style = null
                                },
                                new()
                                {
                                    Value = "Mr. Miyagi",
                                    Style = null
                                }
                            }
                        }
                    }
                }
            }
        };
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        WorkbookWriter writer = new();
        MemoryStream ms = writer.Write(npoiWorkbook);

        string directory = @"C:\\Users\\manue\\OneDrive\\Documents\\Repositories\\test-excel";
        string fileName = "1.xlsx";
        string path = Path.Combine(directory, fileName);
        byte[] fileBytes = ms.ToArray();
        File.WriteAllBytes(path, fileBytes);
    }

    [TestMethod]
    public void TestMethod2()
    {
        Workbook excelWorkbook = new()
        {
            FileType = FileType.Xlsx,
            Sheets = new List<Sheet>()
        };

        for (int z = 0; z < 5; z++)
        {
            Sheet newSheet = new()
            {
                Name = z.ToString(),
                Rows = new List<Row>(),
                Style = new()
                {
                    ColorBanding = new()
                    {
                        EvenRows = Color.Blue,
                        OddRows = Color.Orange
                    }
                }
            };

            for (int i = 0; i < 1000; i++)
            {
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                for (int y = 0; y < 5; y++)
                {
                    Cell newCell = new()
                    {
                        Value = y
                    };
                    newRow.Cells.Add(newCell);
                }

                newSheet.Rows.Add(newRow);
            }

            excelWorkbook.Sheets.Add(newSheet);
        }
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);

        WorkbookWriter writer = new();
        MemoryStream ms = writer.Write(npoiWorkbook);

        string directory = @"C:\\Users\\manue\\OneDrive\\Documents\\Repositories\\test-excel";
        string fileName = "3.xlsx";
        string path = Path.Combine(directory, fileName);
        byte[] fileBytes = ms.ToArray();
        File.WriteAllBytes(path, fileBytes);
    }

    [TestMethod]
    public void TestMethod3()
    {
        Style redBackground = new()
        {
            FillForegroundColor = Color.Red
        };
        Style greenBackground = new()
        {
            FillForegroundColor = Color.Green
        };
        Style blueBackground = new()
        {
            FillForegroundColor = Color.Blue
        };
        Style boldFont = new()
        {
            FontStyle = new()
            {
                IsBold = true
            }
        };

        var rows = new List<Row>();

        for (int i = 0; i < 10; i++)
        {
            List<Cell> cells = Enumerable.Range(0, 10)
                .Select(i => new Cell() { Value = i, Style = i == 0 ? boldFont : null })
                .ToList();

            Style s = (i % 3) switch
            {
                0 => redBackground,
                1 => greenBackground,
                2 => blueBackground,
                _ => null
            };

            Row row = new()
            {
                Cells = cells,
                Style = s
            };
            rows.Add(row);
        }

        Sheet sheet = new()
        {
            Rows = rows
        };

        Workbook excelWorkbook = new()
        {
            FileType = FileType.Xlsx,
            Sheets = new List<Sheet>
            {
                sheet
            }
        };

        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);

        WorkbookWriter writer = new();
        MemoryStream ms = writer.Write(npoiWorkbook);

        string directory = @"C:\\Users\\manue\\OneDrive\\Documents\\Repositories\\test-excel";
        string fileName = "4.xlsx";
        string path = Path.Combine(directory, fileName);
        byte[] fileBytes = ms.ToArray();
        File.WriteAllBytes(path, fileBytes);
    }
}