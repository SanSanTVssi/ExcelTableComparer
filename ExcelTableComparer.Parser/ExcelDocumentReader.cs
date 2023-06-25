using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public class ExcelDocumentReader : IExcelDocumentReader
{
    public ExcelDocumentReader()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public List<List<string>> ReadExcelTable(string excelFile, string listName)
    {
        using var package = new ExcelPackage(new FileInfo(excelFile));
        var workbook = package.Workbook;
        var worksheet = workbook.Worksheets[listName];
        var dataCollection = new List<List<string>>();

        for (var row = 1; row <= worksheet.Dimension.Rows; row++)
        {
            var rowData = new List<string>();
            for (var col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value;
                var cellData = cellValue != null ? cellValue.ToString() : string.Empty;
                rowData.Add(cellData);
            }
            dataCollection.Add(rowData);
        }

        return dataCollection;
    }
}