using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public interface IExcelDocumentReader
{
    ExcelWorksheetSlim ReadDocumentWorksheet(ExcelPackage package, string worksheetName);
    ExcelWorksheetSlim ReadDocumentWorksheet(ExcelWorksheet worksheet);
}