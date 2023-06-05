﻿using ExcelCommander.Base.ClientServer;

namespace ExcelCommander.Services
{
    internal sealed class StandaloneUse
    {
        public ExcelWriter Writer { get; }
        public string OutputFile { get; }
        public StandaloneUse(string outputFile)
        {
            OutputFile = outputFile;
            Writer = new ExcelWriter(outputFile);

            Console.WriteLine($"Write to file {outputFile}.");
        }
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
        }
        public void ExecuteCommand(string command)
        {
            Writer.EvaluateCommand(command);
        }
    }

    internal sealed class SocketUse: IDisposable
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
            Client.Send(new Base.Serialization.CommandData
            {
                CommandType = "Development",
                Contents = command
            });
        }
        #endregion
    }
}
