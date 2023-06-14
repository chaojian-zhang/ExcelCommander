using System;
using System.IO;
using System.Linq;
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

        #region Routines
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

        #region Helper
        public void ConstructPayloads()
        {
            switch (CommandType)
            {
                case "Value bool":
                    Payload = (bool)bool.Parse(Contents);
                    break;
                case "Value int":
                    Payload = (int)int.Parse(Contents);
                    break;
                case "Value double":
                    Payload = (double)double.Parse(Contents);
                    break;
                case "Value char":
                    Payload = (char)Contents.First();
                    break;
                case "Value string":
                    Payload = (string)Contents;
                    break;
                case "Value bool[]":
                    Payload = (bool[])Contents
                        .Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => bool.Parse(v))
                        .ToArray();
                    break;
                case "Value int[]":
                    Payload = (int[])Contents
                        .Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => int.Parse(v))
                        .ToArray();
                    break;
                case "Value double[]":
                    Payload =(double[])Contents
                        .Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => double.Parse(v))
                        .ToArray();
                    break;
                case "Value char[]":
                    Payload = (char[])Contents
                        .Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => v.First())
                        .ToArray();
                    break;
                case "Value string[]":
                    Payload = (string[])Contents.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    break;
                default:
                    break;
            }
        }
        #endregion
    }
}
