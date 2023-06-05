using ExcelCommander.Base.ClientServer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCommander.Services
{
    internal abstract class HandlerBase
    {
        public void Execute(string[] commands, bool interpretIfNull = true)
        {
            if (commands == null && interpretIfNull)
            {
                string input = Console.ReadLine();
                ExecuteCommand(input);
            }
            else
            {
                foreach (var command in commands)
                    ExecuteCommand(command);
            }
        }
        public abstract void ExecuteCommand(string command);
    }

    internal sealed class StandaloneUse : HandlerBase
    {
        public string OutputFile { get; }
        public StandaloneUse(string outputFile)
        {
            OutputFile = outputFile;
        }

        public override void ExecuteCommand(string command)
        {
            throw new NotImplementedException();
        }
    }

    internal sealed class SocketUse: HandlerBase, IDisposable
    {
        #region Construction
        private int Port { get; }
        private Client Client { get; }
        public SocketUse(int port)
        {
            Port = port;

            Client = new Client(port, data => null);
            Client.Start();
        }
        #endregion

        #region Disposal

        #endregion

        #region Handling
        public override void ExecuteCommand(string command)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            Client.Close();
        }
        #endregion
    }
}
