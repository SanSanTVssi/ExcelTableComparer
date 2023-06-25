using System.Collections;
using CommandLine;
using CommandLine.Text;

namespace ExcelTableComparer;

class Options
{
    [Option('f', "filename", Required = true, HelpText = "Path to input file")]
    public string InputFile { get; set; }

    [Option('o', "outputFile", Required = true, HelpText = "Output filename")]
    public string OutputFile { get; set; }

    [Option('l', "tableListNames", Separator = ',')]
    public IEnumerable<string> TableListNames { get; set; }
}
