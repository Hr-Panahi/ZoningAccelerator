using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace ZoningAccelerator.Helpers
{
    public class FileHelper
    {
        public static void WriteToFile(List<string> lines, string filePath)
        {
            File.WriteAllLines(filePath, lines);
        }

        public static void WriteToExcel(string filePath, List<string> col1, List<string> col2, List<string> col3, List<string> col4)
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var ws = workbook.Worksheets.Add("Comparison Results");

            ws.Cell(1, 1).Value = "New Dwelling Types";
            ws.Cell(1, 2).Value = "New Ancillary Types";
            ws.Cell(1, 3).Value = "New Permitted Uses";
            ws.Cell(1, 4).Value = "New TypeOfUses";

            int maxRows = Math.Max(
                Math.Max(col1.Count, col2.Count),
                Math.Max(col3.Count, col4.Count)
            );

            for (int i = 0; i < maxRows; i++)
            {
                if (i < col1.Count) ws.Cell(i + 2, 1).Value = col1[i];
                if (i < col2.Count) ws.Cell(i + 2, 2).Value = col2[i];
                if (i < col3.Count) ws.Cell(i + 2, 3).Value = col3[i];
                if (i < col4.Count) ws.Cell(i + 2, 4).Value = col4[i];
            }

            workbook.SaveAs(filePath);
        }

    }
}
