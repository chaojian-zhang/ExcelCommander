using ExcelCommander.Base;

namespace XlsxCommander
{
    public sealed class XlsxCommander
    {
        #region Constructor
        public ExcelWriter Writer { get; }
        public string OutputFile { get; }

        public static XlsxCommander Start(string outputFile)
            => new XlsxCommander(outputFile);
        public XlsxCommander(string outputFile)
        {
            OutputFile = outputFile;
            Writer = new ExcelWriter(outputFile);

            Console.WriteLine($"Write to file {outputFile}.");
        }
        #endregion

        #region Parsing Methods
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
        #endregion
    }
}
