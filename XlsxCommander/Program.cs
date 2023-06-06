namespace XlsxCommander
{
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
                new XlsxCommander(target).Execute(scriptLines);
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