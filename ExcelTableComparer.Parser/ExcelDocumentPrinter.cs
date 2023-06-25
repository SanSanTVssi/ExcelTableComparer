using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public class ExcelDocumentPrinter : IExcelDocumentPrinter
{
    public void PrintExcelWorksheet(IEnumerable<IEnumerable<string>> table)
    {
        foreach (var rowData in table)
        {
            foreach (var cellData in rowData)
            {
                Console.Write(cellData + "\t");
            }
            Console.WriteLine();
        }
    }

    public void PrintExcelWorksheetToFile(IEnumerable<IEnumerable<string>> table, string fileName)
    {
        foreach (var rowData in table)
        {
            foreach (var cellData in rowData)
            {
                File.AppendAllText(fileName, cellData + "\t");
            }
            File.AppendAllText(fileName, "\n");
        }
    }

    public void SaveWorksheetToExcelDocument(ExcelPackage package, ExcelWorksheetSlim worksheetSlim)
    {
        var worksheet = package.Workbook.Worksheets.Add(worksheetSlim.Name);

        for (var row = 0; row < worksheetSlim.Count; row++)
        {
            var rowData = worksheetSlim[row];
            for (var col = 0; col < rowData.Count; col++)
            {
                worksheet.Cells[row + 1, col + 1].Value = rowData[col];
            }
        }
    }
}