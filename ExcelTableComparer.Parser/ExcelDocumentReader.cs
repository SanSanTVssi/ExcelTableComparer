using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public class ExcelDocumentReader : IExcelDocumentReader
{
    public ExcelDocumentReader()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public ExcelWorksheetSlim ReadDocumentWorksheet(ExcelPackage package, string worksheetName) =>
        ReadDocumentWorksheet(package.Workbook.Worksheets[worksheetName]);

    public ExcelWorksheetSlim ReadDocumentWorksheet(ExcelWorksheet worksheet)
    {
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

        return new ExcelWorksheetSlim(worksheet.Name, dataCollection);
    }

}