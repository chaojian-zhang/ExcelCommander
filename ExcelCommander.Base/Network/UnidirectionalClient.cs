using System;
using System.Net;
using System.Net.Sockets;

namespace ExcelCommander.Base.Network
{
    public class UnidirectionalClient : IDisposable
    {
        #region Config
        public static readonly string HostAddress = "127.0.0.1";
        public const int BufferSize = 64 * 1024 * 1024; // 64 Mb
        #endregion

        #region Lifetime
        public void Dispose()
        {
            try
            {
                Socket.Shutdown(SocketShutdown.Both);
            }
            catch (Exception){}
            finally
            {
                Socket.Dispose();
            }
        }
        #endregion

        #region Members
        Socket Socket;
        #endregion

        #region Entry
        public Socket StartClient(int servicePort)
        {
            IPHostEntry entry = Dns.GetHostEntry(HostAddress);
            IPEndPoint endpoint = new IPEndPoint(entry.AddressList[0], servicePort);
            Socket = new Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            Socket.Connect(endpoint);
            return Socket;
        }
        #endregion

        #region Messaging
        public void Send(Socket connection, byte[] data)
        {
            if (data.Length > BufferSize)
                throw new ArgumentException("Invalid data size.");

            connection.Send(data);
        }
        public void SendAndReceive(Socket connection, byte[] data, out byte[] replyData, out int replyLength)
        {
            if (data.Length > BufferSize)
                throw new ArgumentException("Invalid data size.");

            // Send
            connection.Send(data);
            // Receive
            replyData = new byte[BufferSize];
            replyLength = connection.Receive(replyData);
        }
        #endregion
    }
}