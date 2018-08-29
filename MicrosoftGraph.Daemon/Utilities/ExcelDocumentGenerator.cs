using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace MicrosoftGraph.Daemon.Utilities
{
    public class ExcelDocumentGenerator
    {
        public MemoryStream GenerateDocument()
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Test");

                var excelWorkSheet = excel.Workbook.Worksheets["Test"];

                List<string[]> headerRow = new List<string[]>()
                {
                    new string[]
                    {
                        "Test Column 1",
                        "Test Column 2",
                        "Test Column 3"
                    }
                };

                // Determine the header range (e.g. A1:E1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Popular header row data
                excelWorkSheet.Cells[headerRange].LoadFromArrays(headerRow);

                var cellData = new List<object[]>
                {
                    new object[] {"Test data 1.1", "Test data 2.1", "Test data 3.1",},
                    new object[] {"Test data 1.2", "Test data 2.2", "Test data 3.2",},
                    new object[] {"Test data 1.3", "Test data 2.3", "Test data 3.3",},
                    new object[] {"Test data 1.4", "Test data 2.4", "Test data 3.4",},
                    new object[] {"Test data 1.5", "Test data 2.5", "Test data 3.5",},
                };

                excelWorkSheet.Cells[2, 1].LoadFromArrays(cellData);

                MemoryStream stream = new MemoryStream();
                excel.SaveAs(stream);
                stream.Position = 0;

                return stream;
            }
        }
    }
}
