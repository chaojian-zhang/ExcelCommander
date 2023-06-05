using ExcelCommander.Base.ClientServer;
using ExcelCommander.Base.Serialization;

namespace ExcelCommanderUnitTests
{
    public class NetworkServiceTest
    {
        [Fact]
        public void ClientCanCommunicateWithServer()
        {
            var server = new Server(data =>
            {
                return new CommandData()
                {
                    CommandType = "Reply",
                    Contents = data.Contents
                };
            });
            int port = server.Start();

            var client = new Client(port, data =>
            {
                Assert.Equal("Reply", data.CommandType);
                return null;
            });
            client.Start();

            client.Send(new CommandData()
            {
                CommandType = "test",
                Contents = "Hello world"
            });
        }
    }
}