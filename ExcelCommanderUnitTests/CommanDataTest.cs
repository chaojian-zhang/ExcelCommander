using ExcelCommander.Base.Serialization;

namespace ExcelCommanderUnitTests
{
    public class CommanDataTest
    {
        [Fact]
        public void CompressedSerializationShouldWork()
        {
            var data = new CommandData()
            {
                CommandType = "test",
                Contents = "Hello world"
            };
            var bytes = data.Serialize();
            var dataRecovered = CommandData.Deserialize(bytes, bytes.Length);
        }
    }
}