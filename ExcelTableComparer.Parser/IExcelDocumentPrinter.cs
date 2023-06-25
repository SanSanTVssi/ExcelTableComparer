using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public interface IExcelDocumentPrinter
{
    void PrintExcelWorksheet(IEnumerable<IEnumerable<string>> table);
    void PrintExcelWorksheetToFile(IEnumerable<IEnumerable<string>> table, string fileName);
    void SaveWorksheetToExcelDocument(ExcelPackage package, ExcelWorksheetSlim worksheetSlim);
}