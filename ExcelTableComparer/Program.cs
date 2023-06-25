using CommandLine;
using ExcelTableComparer;
using ExcelTableComparer.Parser;

var Logger = NLog.LogManager.GetCurrentClassLogger();
var parser = new Parser(settings =>
{
    settings.AutoHelp = true;
    settings.AutoVersion = true;
    settings.HelpWriter = Console.Out;
});

var reader = new ExcelDocumentReader();
var printer = new ExcelDocumentPrinter();
var statusCode = 0;

parser
    .ParseArguments<Options>(args)
    .WithParsed(options =>
    {
        try
        {
            var excelFile = options.InputFile;
            var excelListsName = options.TableListNames.ToList();

            var table1 = reader.ReadExcelTable(excelFile, excelListsName[0]);
            var table2 = reader.ReadExcelTable(excelFile, excelListsName[1]);

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
            Logger.Info($"test result: {res}, result1: {result1.Count}, result2: {result2.Count}. Done!");
        }
        catch (Exception ex)
        {
            Logger.Fatal($"Fatal error occurred: {ex.Message}");
            statusCode = 1;
        }
    });

return statusCode;

