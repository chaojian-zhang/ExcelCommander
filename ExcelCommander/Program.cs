using ExcelCommander.Base.ClientServer;
using ExcelCommander.Services;

namespace ExcelCommander
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("""
                    Missing inputs.
                    ExcelCommander <Output Excel Filename.xlsx>|<Server Port Number> (<ScriptFilePath>)
                    """);
                return;
            }

            string target = args.First();
            string[] scriptLines = args.Length >= 2 
                ? File.ReadAllLines(Path.GetFullPath(args[1]))
                : null;

            if (int.TryParse(target, out int port))
            {
                new SocketUse(port).Execute(scriptLines);
            }
            else if (Path.GetExtension(target) == ".xlsx")
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