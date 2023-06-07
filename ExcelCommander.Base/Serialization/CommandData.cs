using System.IO;
using System.Text;

namespace ExcelCommander.Base.Serialization
{
    public class CommandData
    {
        #region Properties
        public string CommandType;
        public string Contents;
        public object Payload; // Remark-cz: Arbitrary payload for runtime calls, not serialized; Only useful for XlsxCommander
        #endregion

        #region Interface
        public byte[] Serialize()
        {
            using (MemoryStream memory = new MemoryStream())
            using (BinaryWriter writer = new BinaryWriter(memory, Encoding.UTF8, false))
            {
                WriteToStream(writer, this);
                return memory.ToArray();
            }
        }
        public static CommandData Deserialize(byte[] data, int length, int offset = 0) // Remark-cz: At the moment when multiple Send is sent, they may come all together in one single Receive (by the server), so we must implement custom breakdown of such messaging
        {
            using (MemoryStream memory = new MemoryStream(data, offset, length))
            using (BinaryReader reader = new BinaryReader(memory, Encoding.UTF8, false))
                return ReadFromStream(reader);
        }
        #endregion

        #region Helpers
        private static void WriteToStream(BinaryWriter writer, CommandData data)
        {
            // Explicitly write string length
            int headerLength = Encoding.UTF8.GetBytes(data.CommandType).Length;
            int contentLength = Encoding.UTF8.GetBytes(data.Contents).Length;
            int frameSize = headerLength + contentLength + sizeof(int) * 3; // Remark: x3 for headerLength, contentLength and frameSize itself
            writer.Write(frameSize);

            writer.Write(headerLength);
            writer.Write(Encoding.UTF8.GetBytes(data.CommandType));
            writer.Write(contentLength);
            writer.Write(Encoding.UTF8.GetBytes(data.Contents));
        }
        private static CommandData ReadFromStream(BinaryReader reader)
        {
            CommandData data = new CommandData();

            reader.ReadInt32(); // Remark-cz: Frame length is not useful for us; It's for receiver
            int headerLength = reader.ReadInt32();
            data.CommandType = Encoding.UTF8.GetString(reader.ReadBytes(headerLength));
            int contentLength = reader.ReadInt32();
            data.Contents = Encoding.UTF8.GetString(reader.ReadBytes(contentLength));

            return data;
        }
        #endregion
    }
}
