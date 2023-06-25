namespace ExcelTableComparer.Parser;

public interface IExcelDocumentPrinter
{
    void PrintExcelTable(IEnumerable<IEnumerable<string>> table);
    void SaveTableToExcelDocument(string excelFile, string listName, List<List<string>> dataCollection);
}