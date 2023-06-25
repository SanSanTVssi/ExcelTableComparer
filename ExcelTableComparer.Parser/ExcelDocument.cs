using OfficeOpenXml;

namespace ExcelTableComparer.Parser;

public class ExcelDocument : IDisposable
{
    private ExcelPackage m_package;
    private List<ExcelWorksheetSlim> m_tables = new();
    private IExcelDocumentReader m_reader;
    private IExcelDocumentPrinter m_printer;

    public string Name { get; }

    public ExcelDocument(string pathToDocument) : this(new ExcelDocumentReader(), new ExcelDocumentPrinter(),
        pathToDocument)
    {
        Name = pathToDocument;
    }

    public ExcelDocument(IExcelDocumentReader reader, IExcelDocumentPrinter printer, string pathToDocument)
    {
        m_reader = reader;
        m_printer = printer;
        m_package = new (pathToDocument);
    }

    private ExcelDocument(ExcelDocument document)
    {
        m_package = document.m_package;
        m_tables = document.m_tables;
        m_reader = document.m_reader;
        m_printer = document.m_printer;
    }

    public ExcelDocument ReadDocument()
    {
        foreach (var worksheet in m_package.Workbook.Worksheets)
        {
            m_tables.Add(m_reader.ReadDocumentWorksheet(worksheet));
        }

        return new ExcelDocument(this);
    }

    public ExcelDocument AddWorksheet(ExcelWorksheetSlim newWorksheet)
    {
        m_printer.SaveWorksheetToExcelDocument(m_package, newWorksheet);
        m_tables.Add(newWorksheet);
        return this;
    }

    public ExcelWorksheetSlim this[int index] => m_tables[index];
    public ExcelWorksheetSlim? this[string index] => m_tables.Find(worksheet => worksheet.Name == index);

    public void SaveDocument()
    {
        m_package.Save();
    }

    public void Dispose()
    {
        m_package.Dispose();
    }
}