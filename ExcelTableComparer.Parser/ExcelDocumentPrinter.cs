using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public class ExcelDocumentPrinter : IExcelDocumentPrinter
{
    public void PrintExcelTable(IEnumerable<IEnumerable<string>> table)
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

    public void PrintExcelTableToFile(IEnumerable<IEnumerable<string>> table, string fileName)
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

    public void SaveTableToExcelDocument(string excelFile, string listName, List<List<string>> dataCollection)
    {
        using var package = new ExcelPackage();
        var c = package.Workbook.Worksheets.Count;
        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(listName);

        for (int row = 0; row < dataCollection.Count; row++)
        {
            List<string> rowData = dataCollection[row];
            for (int col = 0; col < rowData.Count; col++)
            {
                worksheet.Cells[row + 1, col + 1].Value = rowData[col];
            }
        }

        FileInfo file = new FileInfo(excelFile);
        package.SaveAs(file);
    }
}