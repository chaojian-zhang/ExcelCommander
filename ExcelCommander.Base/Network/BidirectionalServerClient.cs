using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace ExcelCommander.Base.Network
{
    public static class TcpHelper
    {
        public static int FindAvailablePort()
        {
            TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }
    }
    public class BidirectionalServerClient : IDisposable
    {
        #region Config
        public static readonly string HostAddress = "127.0.0.1";
        public const int BufferSize = 8 * 1024 * 1024; // 8 Mb
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
        public int StartServer(Action<int, byte[], Socket> callback)
        {
            List<Socket> clients = new List<Socket>();

            int servicePort = TcpHelper.FindAvailablePort();
            IPHostEntry entry = Dns.GetHostEntry(HostAddress);
            IPEndPoint endpoint = new IPEndPoint(entry.AddressList[0], servicePort);
            Socket = new Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            Socket.Bind(endpoint);
            Socket.Listen(100);
            new Thread(() =>
            {
                try
                {
                    while (true)
                    {
                        var client = Socket.Accept();
                        Console.WriteLine("New client is connected.");
                        clients.Add(client);
                        new Thread(() => ServerHandleClient(client)).Start();
                    }
                }
                catch (Exception){}
            }).Start();
            return servicePort;

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
                    Console.WriteLine(e.Message);
                }
            }
        }
        public Socket StartClient(int servicePort, Action<int, byte[]> callback)
        {
            IPHostEntry entry = Dns.GetHostEntry(HostAddress);
            IPEndPoint endpoint = new IPEndPoint(entry.AddressList[0], servicePort);
            Socket = new Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            Socket.Connect(endpoint);
            new Thread(() => ClientReceiveMessage(Socket)).Start();
            return Socket;

            void ClientReceiveMessage(Socket socket)
            {
                try
                {
                    while (true)
                    {
                        byte[] buffer = new byte[BufferSize];
                        var size = socket.Receive(buffer);
                        callback(size, buffer);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
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