namespace ExcelTableComparer.Parser;

public class ExcelWorksheetSlim : List<List<string>>
{
    public string Name { get; }

    public ExcelWorksheetSlim(string name, List<List<string>> worksheet) : base(worksheet)
    {
        Name = name;
    }
}