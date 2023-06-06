using ExcelCommander.Base.ClientServer;
using ExcelCommander.Base.Serialization;
using ExcelCommander.Base;

namespace ExcelCommander
{
    public sealed class ExcelCommander : IDisposable
    {
        #region Construction
        private int Port { get; }
        private Client Client { get; set; }
        public static ExcelCommander Connect(int port)
            => new ExcelCommander(port);
        public ExcelCommander(int port)
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
}
