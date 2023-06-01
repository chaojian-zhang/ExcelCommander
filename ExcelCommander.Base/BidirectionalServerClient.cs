using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Console = Colorful.Console;

namespace ExcelCommander.Base
{
    public class BidirectionalServerClient : IDisposable
    {
        #region Config
        public static readonly string HostAddress = "127.0.0.1";
        public const int ServicePort = 12900; // TODO: Remark-cz: Automatically find a new port and report it, so we can interface with multiple excel instances
        public const int BufferSize = 64 * 1024 * 1024; // 64 Mb
        #endregion

        #region Lifetime
        public void Dispose()
        {
            Socket.Shutdown(SocketShutdown.Both);
        }
        #endregion

        #region Members
        Socket Socket;
        #endregion

        #region Entry
        public void StartServer(Action<int, byte[], Socket> callback)
        {
            List<Socket> clients = new List<Socket>();

            IPHostEntry entry = Dns.GetHostEntry(HostAddress);
            IPEndPoint endpoint = new IPEndPoint(entry.AddressList[0], ServicePort);
            Socket = new Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            Socket.Bind(endpoint);
            Socket.Listen(100);
            new Thread(() =>
            {
                while (true)
                {
                    var client = Socket.Accept();
                    Console.WriteLine("New client is connected.");
                    clients.Add(client);
                    new Thread(() => ServerHandleClient(client)).Start();
                }
            }).Start();

            void ServerHandleClient(Socket client)
            {
                try
                {
                    while (true)
                    {
                        byte[] buffer = new byte[BufferSize];
                        var size = client.Receive(buffer);
                        callback(size, buffer, client);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message, Color.Red);
                }
            }
        }
        public Socket StartClient(Action<int, byte[]> callback)
        {
            IPHostEntry entry = Dns.GetHostEntry(HostAddress);
            IPEndPoint endpoint = new IPEndPoint(entry.AddressList[0], ServicePort);
            Socket = new Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            Socket.Connect(endpoint);
            new Thread(() => ClientReceiveMessage(Socket)).Start();
            return Socket;

            void ClientReceiveMessage(Socket socket)
            {
                while (true)
                {
                    byte[] buffer = new byte[BufferSize];
                    var size = socket.Receive(buffer);
                    callback(size, buffer);
                }
            }
        }
        #endregion

        #region Messaging
        public void Send(Socket connection, byte[] data)
        {
            if (data.Length > BufferSize)
                throw new ArgumentException("Invalid data size.");

            connection.Send(data);
        }
        #endregion
    }
}