using K4os.Compression.LZ4.Streams;
using System.IO;
using System.Text;

namespace ExcelCommander.Base.Serialization
{
    public class CommandData
    {
        #region Properties
        public string CommandType;
        public string Contents;
        #endregion

        #region Interface
        public byte[] Serialize()
        {
            using (MemoryStream memory = new MemoryStream())
            using (LZ4EncoderStream stream = LZ4Stream.Encode(memory))
            using (BinaryWriter writer = new BinaryWriter(stream, Encoding.UTF8, false))
            {
                WriteToStream(writer, this);
                return memory.ToArray();
            }
        }
        public static CommandData Deserialize(byte[] data, int length)
        {
            using (MemoryStream memory = new MemoryStream(data, 0, length))
            using (LZ4DecoderStream source = LZ4Stream.Decode(memory))
            using (BinaryReader reader = new BinaryReader(memory, Encoding.UTF8, false))
                return ReadFromStream(reader);
        }
        #endregion

        #region Helpers
        private static void WriteToStream(BinaryWriter writer, CommandData data)
        {
            writer.Write(data.CommandType);
            writer.Write(data.Contents);
        }
        private static CommandData ReadFromStream(BinaryReader reader)
        {
            CommandData data = new CommandData();

            data.CommandType = reader.ReadString();
            data.Contents = reader.ReadString();

            return data;
        }
        #endregion
    }
}
