using ExcelCommander.Base.Network;
using System.Net.Sockets;
using System;
using ExcelCommander.Base.Serialization;
using System.Linq;

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
            ServicePort = Service.StartServer((length, data, client) =>
            {
                // Deal with multiple frames
                int remainingSize = length;
                while (remainingSize > 0)
                {
                    int startIndex = length - remainingSize;
                    int frameSize = BitConverter.ToInt32(data, startIndex);
                    Callback(frameSize, data, startIndex, client);
                    remainingSize -= frameSize;
                    if (remainingSize < 0)
                        throw new ApplicationException("Frame size error.");
                }
            });
            return ServicePort;
        }
        public void Stop()
        {
            Service.Dispose();
        }
        #endregion

        #region Data Marshal
        private void Callback(int length, byte[] data, int offset, Socket client)
        {
            try
            {
                CommandData command = CommandData.Deserialize(data, length, offset);
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
