namespace ExcelTableComparer.Parser;

public interface IExcelDocumentReader
{
    List<List<string>> ReadExcelTable(string excelFile, string listName);
}