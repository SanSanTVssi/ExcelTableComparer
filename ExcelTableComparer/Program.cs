using System.Collections;
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

var statusCode = 0;

parser
    .ParseArguments<Options>(args)
    .WithParsed(options =>
    {
        try
        {
            var input = options.InputFile;

            var excelDocument = new ExcelDocument(input);
            excelDocument.ReadDocument();

            Logger.Trace($"Readed {excelDocument[0].Count} rows from first worksheet");
            Logger.Trace($"Readed {excelDocument[1].Count} rows from second worksheet");

            bool Predicate(List<string> row1, List<string> row2) => row1[3] == row2[3] && row1[2] == row2[2];

            // TODO: It is necessary to add the ability to create custom predicates
            bool ValidateRows(List<string> row1, List<string> row2) => row1.Count > 3 && row2.Count > 3;

            Func<List<List<string>>, List<List<string>>, List<List<string>>> functor
                = (t1, t2) => t2
                    .Where(row2 =>
                        !t1.Any(row1 => ValidateRows(row1, row2) && Predicate(row1, row2))
                    )
                    .ToList();

            Func<List<List<string>>, List<List<string>>, List<List<string>>> test
                = (t1, t2) => t2
                    .Where(row2 =>
                        !t1.Any(row1 => ValidateRows(row1, row2) && !Predicate(row1, row2))
                    )
                    .ToList();

            var result1 = functor(excelDocument[1], excelDocument[0]);
            Logger.Trace($"{result1.Count} rows from second table");

            var result2 = functor(excelDocument[0], excelDocument[1]);
            Logger.Trace($"{result2.Count} rows from first table");


            var result1Worksheet = new ExcelWorksheetSlim(excelDocument[0].Name, result1);
            var result2Worksheet = new ExcelWorksheetSlim(excelDocument[1].Name, result2);

            var res = test(result1, result2).Count;
            Logger.Trace($"Test result: {res}");
            if (res != 0)
            {
                Logger.Error("Test failed");
                statusCode = 1;
            }

            File.Delete(options.OutputFile);
            var output = new ExcelDocument(options.OutputFile);
            output.AddWorksheet(result1Worksheet);
            output.AddWorksheet(result2Worksheet);
            output.SaveDocument();
        }
        catch (Exception ex)
        {
            Logger.Fatal($"Fatal error occurred: {ex.Message}");
            statusCode = 1;
        }
    });

return statusCode;

