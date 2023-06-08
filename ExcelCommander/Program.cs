using System.Net.Sockets;

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
                    ExcelCommander <Server Port Number> (<ScriptFilePath>)
                    """);
                return;
            }

            string target = args.First();
            string[] scriptLines = args.Length >= 2 
                ? File.ReadAllLines(Path.GetFullPath(args[1]))
                : null;

            if (int.TryParse(target, out int port))
            {
                try
                {
                    var commander = new ExcelCommander(port);
                    commander.Execute(scriptLines);
                    commander.Dispose();
                }
                catch (SocketException)
                {
                    Console.WriteLine("Cannot connect to service. Check and make sure service is online and port number is correct.");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error: {e}");
                }
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