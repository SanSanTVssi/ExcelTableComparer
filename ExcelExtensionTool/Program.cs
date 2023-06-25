using System.Collections;
using ExcelExtensionTool;

try
{
    var excelFile = "/Users/aai/Desktop/in.xlsx";
    var reared = new ExcelDocumentReader();
    var printer = new ExcelDocumentPrinter();
    var table1 = reared.ReadExcelTable(excelFile, "Data0");
    var table2 = reared.ReadExcelTable(excelFile, "Data1");

    Func<List<List<string>>, List<List<string>>, List<List<string>>> functor
        = (t1, t2) => t2
            .Where(row2 =>
                !t1.Any(row1 => row1.Count > 3 && row1[3] == row2[3] && row1[2] == row2[2])
            )
            .ToList();

    Func<List<List<string>>, List<List<string>>, List<List<string>>> test
        = (t1, t2) => t2
            .Where(row2 =>
                !t1.Any(row1 => row1.Count > 3 && !(row1[3] == row2[3] && row1[2] == row2[2]))
            )
            .ToList();

    var result2 = functor(table1, table2);
    var result1 = functor(table2, table1);

    var output1 = "/Users/aai/Desktop/output1.xlsx";
    var output2 = "/Users/aai/Desktop/output2.xlsx";

    File.Delete(output1);
    File.Delete(output2);

    var res = test(result1, result2).Count;
    printer.SaveTableToExcelDocument(output1,"List1", result1);
    printer.SaveTableToExcelDocument(output2,"List2", result2);
    Console.WriteLine($"test result: {res}, result1: {result1.Count}, result2: {result2.Count}. Done!");
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex);
    return 1;
}

return 0;

