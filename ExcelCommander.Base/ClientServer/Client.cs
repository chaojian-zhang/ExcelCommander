using ExcelCommander.Base.Network;
using System.Net.Sockets;
using System;
using ExcelCommander.Base.Serialization;

namespace ExcelCommander.Base.ClientServer
{
    public class Client
    {
        #region Internal Data
        private BidirectionalServerClient ClientInstance;
        private Socket ServerReference;
        private Func<CommandData, CommandData> CommandHandler;
        private int ServicePort;
        #endregion

        #region Constructor
        public Client(int servicePort, Func<CommandData, CommandData> handler)
        {
            ServicePort = servicePort;
            CommandHandler = handler;
        }
        #endregion

        #region Method
        public void Start()
        {
            ClientInstance = new BidirectionalServerClient();
            ServerReference = ClientInstance.StartClient(ServicePort, (length, data) => Callback(length, data));
        }
        public void Send(CommandData data)
        {
            ClientInstance.Send(ServerReference, data.Serialize());
        }
        #endregion

        #region Data Marshal
        private void Callback(int length, byte[] data)
        {
            try
            {
                CommandData reply = CommandData.Deserialize(data, length);
                CommandHandler?.Invoke(reply);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
        }
        #endregion
    }
}
