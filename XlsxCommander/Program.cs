using ExcelCommander.Base;

namespace XlsxCommander
{
    internal sealed class StandaloneUse
    {
        public ExcelWriter Writer { get; }
        public string OutputFile { get; }
        public StandaloneUse(string outputFile)
        {
            OutputFile = outputFile;
            Writer = new ExcelWriter(outputFile);

            Console.WriteLine($"Write to file {outputFile}.");
        }
        public void Execute(string[] commands, bool interpretIfNull = true)
        {
            if (commands == null && interpretIfNull)
            {
                while (true)
                {
                    Console.Write("> ");
                    string input = Console.ReadLine();
                    ExecuteCommand(input);
                }
            }
            else
            {
                foreach (var command in commands)
                    ExecuteCommand(command);
            }
        }
        public void ExecuteCommand(string command)
        {
            if (command == "Help")
                Console.WriteLine(CommanderHelper.GetHelpString());
            else 
                Writer.EvaluateCommand(command);
        }
    }
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("""
                    Missing inputs.
                    XlsxCommander <Output Excel Filename.xlsx> (<ScriptFilePath>)
                    """);
                return;
            }

            string target = args.First();
            string[] scriptLines = args.Length >= 2
                ? File.ReadAllLines(Path.GetFullPath(args[1]))
                : null;

            if (Path.GetExtension(target) == ".xlsx")
            {
                new StandaloneUse(target).Execute(scriptLines);
            }
            else
            {
                Console.WriteLine($"""
                    Invalid argument: {target}
                    """);
                return;
            }
        }
    }
}