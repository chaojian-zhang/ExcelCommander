using ExcelCommander.Base;
using ExcelCommander.Base.ClientServer;
using ExcelCommander.Base.Serialization;

namespace ExcelCommander
{
    internal sealed class SocketUse : IDisposable
    {
        #region Construction
        private int Port { get; }
        private Client Client { get; set; }
        public SocketUse(int port)
        {
            Port = port;

            Client = new Client(Port, data => null);
            Client.Start();

            Console.WriteLine($"Service started at port {Port}.");
        }
        #endregion

        #region Disposal
        public void Dispose()
        {
            Client.Close();
        }
        #endregion

        #region Handling
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
            Dispose();
        }
        public void ExecuteCommand(string command)
        {
            if (command == "Help")
                Console.WriteLine(CommanderHelper.GetHelpString());
            else if (command.StartsWith("Get"))
            {
                CommandData reply = Client.SendAndReceive(new CommandData
                {
                    CommandType = "Development",
                    Contents = command
                });
                Console.WriteLine($"[{reply.CommandType}] {reply.Contents}");
            }
            else
            {
                Client.Send(new CommandData
                {
                    CommandType = "Development",
                    Contents = command
                });
            }
        }
        #endregion
    }
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
                    new SocketUse(port).Execute(scriptLines);
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