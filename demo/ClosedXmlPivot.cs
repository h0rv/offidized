using ClosedXML.Excel;
using System;

class Program
{
    static void Main()
    {
        var wb = new XLWorkbook();

        // Create data sheet
        var dataSheet = wb.Worksheets.Add("Data");

        // Add headers
        dataSheet.Cell("A1").Value = "Region";
        dataSheet.Cell("B1").Value = "Product";
        dataSheet.Cell("C1").Value = "Quarter";
        dataSheet.Cell("D1").Value = "Sales";

        // Add data
        var data = new object[][]
        {
            new object[] { "North", "Widget", "Q1", 1000 },
            new object[] { "North", "Widget", "Q2", 1500 },
            new object[] { "North", "Gadget", "Q1", 2000 },
            new object[] { "North", "Gadget", "Q2", 2500 },
            new object[] { "South", "Widget", "Q1", 800 },
            new object[] { "South", "Widget", "Q2", 1200 },
            new object[] { "South", "Gadget", "Q1", 1800 },
            new object[] { "South", "Gadget", "Q2", 2200 },
            new object[] { "East", "Widget", "Q1", 1100 },
            new object[] { "East", "Widget", "Q2", 1600 },
            new object[] { "East", "Gadget", "Q1", 2100 },
            new object[] { "East", "Gadget", "Q2", 2600 },
            new object[] { "West", "Widget", "Q1", 900 },
            new object[] { "West", "Widget", "Q2", 1300 },
            new object[] { "West", "Gadget", "Q1", 1900 },
            new object[] { "West", "Gadget", "Q2", 2300 }
        };

        for (int i = 0; i < data.Length; i++)
        {
            dataSheet.Cell(i + 2, 1).Value = (string)data[i][0];
            dataSheet.Cell(i + 2, 2).Value = (string)data[i][1];
            dataSheet.Cell(i + 2, 3).Value = (string)data[i][2];
            dataSheet.Cell(i + 2, 4).Value = (int)data[i][3];
        }

        // Create pivot sheet
        var pivotSheet = wb.Worksheets.Add("Pivot Analysis");

        // Create pivot table
        var dataRange = dataSheet.Range("A1:D17");
        var pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet.Cell("A3"), dataRange);

        // Add row field (Region)
        pivotTable.RowLabels.Add("Region");

        // Add column field (Product)
        pivotTable.ColumnLabels.Add("Product");

        // Add data field (Sum of Sales)
        pivotTable.Values.Add("Sales", "Sum of Sales");

        // Save
        wb.SaveAs("closedxml_pivot.xlsx");
        Console.WriteLine("Created closedxml_pivot.xlsx");
    }
}
