using ExcelCommander.Base.Network;
using System.Net.Sockets;
using System;
using ExcelCommander.Base.Serialization;

namespace ExcelCommander.Base.ClientServer
{
    public class Client
    {
        #region Internal Data
        private UnidirectionalClient ClientInstance;
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
            ClientInstance = new UnidirectionalClient();
            ServerReference = ClientInstance.StartClient(ServicePort);
        }
        public void Send(CommandData data)
        {
            ClientInstance.Send(ServerReference, data.Serialize());
        }
        public CommandData SendAndReceive(CommandData data)
        {
            ClientInstance.SendAndReceive(ServerReference, data.Serialize(), out byte[] replyData, out int replyLength);
            return CommandData.Deserialize(replyData, replyLength);
        }
        public void Close()
        {
            ServerReference.Close();
        }
        #endregion
    }
}
