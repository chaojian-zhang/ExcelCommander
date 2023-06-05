using ExcelCommander.Base.Network;
using System.Net.Sockets;
using System;
using ExcelCommander.Base.Serialization;

namespace ExcelCommander.Base.ClientServer
{
    public class Server
    {
        #region Internal Data
        private BidirectionalServerClient Service;
        public int ServicePort { get; private set; }
        private Func<CommandData, CommandData> CommandHandler;
        #endregion

        #region Constructor
        public Server(Func<CommandData, CommandData> handler)
            => CommandHandler = handler;
        #endregion

        #region Method
        public int Start()
        {
            Service = new BidirectionalServerClient();
            ServicePort = Service.StartServer((length, data, client) => Callback(length, data, client));
            return ServicePort;
        }
        #endregion

        #region Data Marshal
        private void Callback(int length, byte[] data, Socket client)
        {
            try
            {
                CommandData command = CommandData.Deserialize(data, length);
                CommandData reply = CommandHandler?.Invoke(command);
                if (reply != null)
                    Service.Send(client, reply.Serialize());
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
        }
        #endregion
    }
}
