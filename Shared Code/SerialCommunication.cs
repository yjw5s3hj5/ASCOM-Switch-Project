// Author:		Pang Bin (PB) <1371951316@qq.com>
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ASCOM.LocalServer;

namespace SerialCommunication
{
    public static class SerialCommunication
    {
        private static readonly byte[] HEAD = { 0xD9, 0xAB };
        private const byte LOCAL_ADDRESS = 0x00;
        public const byte TAIL = 0xFE;
        private const byte PADDING = 0x00;
        private static readonly int MIN_FRAME_SIZE = 13;
        private static readonly uint[] Crc32Table = new uint[256];

        public static class Cmd1_Driver
        {
            public const byte OTA = 0x00;
            public const byte READ = 0x01;
            public const byte SET_SWITCH = 0x02;
            public const byte SET_VOLTAGE = 0x03;
            public const byte SET_CURRENT = 0x04;
            public const byte GET_INIT_VALUE = 0x05;
            public const byte CLEAR_OTA_FLAG = 0x06;
            public const byte PING = 0x7E;
            public const byte RESET = 0x7F;
        }

        public static class Cmd2_Driver
        {
            // OTA相关命令
            public const byte OTA_START = 0x00;
            public const byte OTA_PAGE = 0x01;
            public const byte OTA_LAST_PAGE = 0x02;
            public const byte OTA_IF_READY = 0x03;
            public const byte OTA_RESET_NOW = 0x04;
            public const byte OTA_CANCEL = 0x05;

            // 开关控制命令
            public const byte SET_SWITCH_ALL_ON = 0x10;
            public const byte SET_SWITCH_ALL_OFF = 0x11;
            // 读写通道选择命令
            public const byte SWITCH_ALL = 0x20;
            public const byte CHANNEL_A1 = 0x21;
            public const byte CHANNEL_A2 = 0x22;
            public const byte CHANNEL_B1 = 0x23;
            public const byte CHANNEL_B2 = 0x24;

            // 传感器读取命令
            public const byte SENSORS = 0x30;
            public const byte BOARD_POWER = 0x31;

            // 空命令
            public const byte CMD_NULL = 0x7F;
        }

        public static class Cmd1_Device
        {
            public const byte RUN_OK = 0x80;
            public const byte ERROR = 0x81;
            public const byte REPORT = 0x82;
            public const byte ALERT = 0x83;
            public const byte PING = 0xFE;
            public const byte STARTUP = 0xFE;
        }

        public static class Cmd2_Device
        {
            public const byte OTA_START = 0x80;
            public const byte OTA_CANCEL = 0x81;
            public const byte OTA_PAGE = 0x82;
            public const byte OTA_READY_RESET = 0x83;
            public const byte OTA_SUCCESS = 0x84;
            public const byte OTA_ERROR_CRC = 0x85;
            public const byte OTA_ERROR_SIZE = 0x86;
            public const byte OTA_WRITE_TWICE = 0x87;

            public const byte CHANNEL_ON = 0x90;
            public const byte CHANNEL_OFF = 0x91;

            public const byte ERROR_UNKNOWN_CMD = 0xFE;
            public const byte CMD_NULL = 0xFF;
        }

        static SerialCommunication()
        {
            // 初始化CRC32 MPEG-2表
            const uint polynomial = 0x04C11DB7;
            for (uint i = 0; i < 256; i++)
            {
                uint crc = i << 24;
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x80000000) != 0)
                        crc = (crc << 1) ^ polynomial;
                    else
                        crc <<= 1;
                }
                Crc32Table[i] = crc;
            }
        }

        public static uint Crc32Mpeg2(byte[] data)
        {
            uint crc = 0xFFFFFFFF;

            foreach (byte b in data)
            {
                byte index = (byte)(((crc >> 24) & 0xFF) ^ b);
                crc = (crc << 8) ^ Crc32Table[index];
            }

            return crc;
        }

        public static byte[] BuildFrame(byte cmd1, byte cmd2, byte[] data)
        {
            List<byte> outData = new List<byte>();
            outData.AddRange(HEAD);
            outData.Add(LOCAL_ADDRESS);
            outData.Add(cmd1);
            outData.Add(cmd2);

            // Add length (little-endian)
            int dataLength = 0;
            if (data != null)
            {
                dataLength = data.Length;
            }
            outData.Add((byte)(dataLength % 256));
            outData.Add((byte)(dataLength / 256));

            outData.AddRange(data);

            // Pad out_data with 0x00 so its length is a multiple of 4
            int padLen = (4 - (outData.Count % 4)) % 4;
            for (int i = 0; i < padLen; i++)
            {
                outData.Add(PADDING);
            }

            // Reverse every 4 bytes in out_data
            List<byte> reversedOutData = new List<byte>();
            for (int i = 0; i < outData.Count; i += 4)
            {
                byte[] chunk = outData.Skip(i).Take(4).Reverse().ToArray();
                reversedOutData.AddRange(chunk);
            }

            // Calculate CRC32
            uint crc32Value = Crc32Mpeg2(reversedOutData.ToArray());
            byte[] crc32Bytes = BitConverter.GetBytes(crc32Value);

            // Remove padding bytes before adding CRC
            List<byte> finalData = outData.Take(outData.Count - padLen).ToList();
            finalData.AddRange(crc32Bytes);

            List<byte> packed_part1 = BytePacking.RepackByteList(finalData.GetRange(3, 4));
            List<byte> packed_part2 = BytePacking.RepackByteList(finalData.GetRange(7, finalData.Count - 7));

            List<byte> finalDataPacked = new List<byte>();
            finalDataPacked.AddRange(HEAD);
            finalDataPacked.Add(LOCAL_ADDRESS);
            finalDataPacked.AddRange(packed_part1);
            finalDataPacked.AddRange(packed_part2);
            finalDataPacked.Add(TAIL); // tail
            return finalDataPacked.ToArray();
        }

        public static FrameParseResult ParseFrame(byte[] receivedFrame)
        {
            var result = new FrameParseResult { IsValid = false };

            // 基本长度检查
            if (receivedFrame == null || receivedFrame.Length < MIN_FRAME_SIZE)
            {
                result.ErrorMessage = $"Frame too short: {receivedFrame?.Length ?? 0} bytes";
                return result;
            }

            int frameLength = receivedFrame.Length;

            // 检查尾字节
            if (receivedFrame[frameLength - 1] != TAIL)
            {
                result.ErrorMessage = $"Invalid Tail Byte: 0x{receivedFrame[frameLength - 1]:X2}";
                return result;
            }

            // 检查头字节
            for (int i = 0; i < HEAD.Length; i++)
            {
                if (receivedFrame[i] != HEAD[i])
                {
                    result.ErrorMessage = $"Invalid Head Byte at position {i}: 0x{receivedFrame[i]:X2}";
                    return result;
                }
            }

            try
            {
                // 提取和解包数据部分 - 使用数组操作而不是LINQ
                var part1Array = new byte[5];
                Array.Copy(receivedFrame, 3, part1Array, 0, 5);
                var unpackedDataPart1 = BytePacking.UnpackByteList(new List<byte>(part1Array));

                var part2Length = frameLength - 9;
                var part2Array = new byte[part2Length];
                Array.Copy(receivedFrame, 8, part2Array, 0, part2Length);
                var unpackedDataPart2 = BytePacking.UnpackByteList(new List<byte>(part2Array));

                // 计算解包后帧的长度
                int unpackedLength = HEAD.Length + 1 + unpackedDataPart1.Count + unpackedDataPart2.Count + 1;
                var receivedFrameUnpacked = new byte[unpackedLength];

                // 构建解包后的帧数组
                int offset = 0;

                // 复制HEAD
                for (int i = 0; i < HEAD.Length; i++)
                {
                    receivedFrameUnpacked[offset++] = HEAD[i];
                }

                receivedFrameUnpacked[offset++] = LOCAL_ADDRESS;

                // 复制解包数据1
                for (int i = 0; i < unpackedDataPart1.Count; i++)
                {
                    receivedFrameUnpacked[offset++] = unpackedDataPart1[i];
                }

                // 复制解包数据2
                for (int i = 0; i < unpackedDataPart2.Count; i++)
                {
                    receivedFrameUnpacked[offset++] = unpackedDataPart2[i];
                }

                receivedFrameUnpacked[offset] = TAIL;

                // 解析命令字节
                byte cmd1 = receivedFrameUnpacked[3];
                byte cmd2 = receivedFrameUnpacked[4];

                // 解析数据长度（小端序）
                int dataLength = receivedFrameUnpacked[5] + (receivedFrameUnpacked[6] << 8);

                // 验证数据长度
                int expectedFrameSize = HEAD.Length + 1 + 2 + 2 + dataLength + 4 + 1;
                if (unpackedLength != expectedFrameSize)
                {
                    result.ErrorMessage = $"Unexpected Frame Size: expected {expectedFrameSize}, got {unpackedLength}";
                    return result;
                }

                // 提取数据（不包含填充）
                byte[] data = new byte[dataLength];
                Array.Copy(receivedFrameUnpacked, 7, data, 0, dataLength);

                // 提取CRC（帧末尾之前的4个字节）
                int crcStart = unpackedLength - 5;
                uint receivedCrc = BitConverter.ToUInt32(receivedFrameUnpacked, crcStart);

                // 为CRC验证构建数据（包括填充）
                int dataForCrcLength = HEAD.Length + 1 + 2 + 2 + dataLength;
                int paddingLength = (4 - (dataForCrcLength % 4)) % 4;
                int totalCrcLength = dataForCrcLength + paddingLength;

                var dataForCrc = new byte[totalCrcLength];

                // 复制前7个字节（头+地址+命令+长度）
                for (int i = 0; i < 7; i++)
                {
                    dataForCrc[i] = receivedFrameUnpacked[i];
                }

                // 复制数据
                for (int i = 0; i < dataLength; i++)
                {
                    dataForCrc[7 + i] = data[i];
                }

                // 填充0
                for (int i = dataForCrcLength; i < totalCrcLength; i++)
                {
                    dataForCrc[i] = 0x00;
                }

                // 每4字节反转（与发送端一致）
                for (int i = 0; i < totalCrcLength; i += 4)
                {
                    // 手动反转4字节块
                    byte temp = dataForCrc[i];
                    dataForCrc[i] = dataForCrc[i + 3];
                    dataForCrc[i + 3] = temp;

                    temp = dataForCrc[i + 1];
                    dataForCrc[i + 1] = dataForCrc[i + 2];
                    dataForCrc[i + 2] = temp;
                }

                // 计算CRC
                uint calculatedCrc = Crc32Mpeg2(dataForCrc);

                // 验证CRC
                if (calculatedCrc != receivedCrc)
                {
                    result.ErrorMessage = $"CRC Mismatch: calculated 0x{calculatedCrc:X8}, received 0x{receivedCrc:X8}";
                    return result;
                }

                // 解析成功
                result.IsValid = true;
                result.Cmd1 = cmd1;
                result.Cmd2 = cmd2;
                result.Data = data;
                result.DataLength = dataLength;

                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"Parsing Exception: {ex.Message}";
                return result;
            }
        }
        public class FrameParseResult
        {
            public bool IsValid { get; set; }
            public byte Cmd1 { get; set; }
            public byte Cmd2 { get; set; }
            public int DataLength { get; set; }
            public byte[] Data { get; set; }
            public string ErrorMessage { get; set; }
        }

        public class BytePacking
        {
            public static List<byte> RepackByteList(List<byte> byteList)
            {
                // 将所有字节转换为二进制字符串
                StringBuilder bitStream = new StringBuilder();
                foreach (byte b in byteList)
                {
                    bitStream.Append(Convert.ToString(b, 2).PadLeft(8, '0'));
                }

                // 每7位一组重新打包
                List<byte> repacked = new List<byte>();
                for (int i = 0; i < bitStream.Length; i += 7)
                {
                    int chunkLength = Math.Min(7, bitStream.Length - i);
                    string chunk = bitStream.ToString(i, chunkLength);

                    // 如果不足7位，用0填充
                    if (chunk.Length < 7)
                    {
                        chunk = chunk.PadRight(7, '0');
                    }

                    // 在最高位添加0，然后转换为字节
                    chunk = "0" + chunk;
                    repacked.Add(Convert.ToByte(chunk, 2));
                }

                return repacked;
            }

            public static List<byte> UnpackByteList(List<byte> repackedList)
            {
                // 将重新打包的字节转换为二进制字符串
                StringBuilder bitStream = new StringBuilder();
                foreach (byte b in repackedList)
                {
                    string bits = Convert.ToString(b, 2).PadLeft(8, '0');
                    // 去掉最高位的0（填充位）
                    bitStream.Append(bits.Substring(1));
                }

                // 每8位一组恢复原始字节
                List<byte> byteList = new List<byte>();
                int originalByteLength = bitStream.Length / 8;

                for (int i = 0; i < originalByteLength * 8; i += 8)
                {
                    int chunkLength = Math.Min(8, bitStream.Length - i);
                    string chunk = bitStream.ToString(i, chunkLength);

                    // 如果不足8位，用0填充
                    if (chunk.Length < 8)
                    {
                        chunk = chunk.PadRight(8, '0');
                    }

                    byteList.Add(Convert.ToByte(chunk, 2));
                }

                return byteList;
            }
        }

        #region Overload methods
        private static FrameParseResult SendMessage(byte cmd1, byte cmd2)
        {
            return SharedResources.SendMessage(cmd1, cmd2, new byte[] { });
        }
        private static FrameParseResult SendMessage(byte cmd1, byte cmd2, List<byte> byte_list)
        {
            return SharedResources.SendMessage(cmd1, cmd2, byte_list.ToArray());
        }
        private static FrameParseResult SendMessage(byte cmd1, byte cmd2, byte value)
        {
            return SharedResources.SendMessage(cmd1, cmd2, new byte[] { value });
        }
        private static FrameParseResult SendMessage(byte cmd1, byte cmd2, uint value)
        {
            return SharedResources.SendMessage(cmd1, cmd2, new byte[] { (byte)(value & 0xFF), (byte)((value >> 8) & 0xFF), (byte)((value >> 16) & 0xFF), (byte)((value >> 24) & 0xFF) });
        }
        private static FrameParseResult SendMessage(byte cmd1, byte cmd2, double value)
        {
            byte[] doubleBytes = BitConverter.GetBytes(value);
            return SharedResources.SendMessage(cmd1, cmd2, doubleBytes);
        }
        private static FrameParseResult SendMessage(byte cmd1, byte cmd2, double[] double_list)
        {
            byte[] data = new byte[double_list.Length * 8];
            for (int i = 0; i < double_list.Length; i++)
            {
                byte[] doubleBytes = BitConverter.GetBytes(double_list[i]);
                Array.Copy(doubleBytes, 0, data, i * 8, 8);
            }
            return SharedResources.SendMessage(cmd1, cmd2, data);
        }
        #endregion 

        #region Switch Commands
        public static FrameParseResult SendMessage(string message)
        {
            switch (message)
            {
                case "TurnOnAllChannels":
                    return SendMessage(Cmd1_Driver.SET_SWITCH, Cmd2_Driver.SET_SWITCH_ALL_ON);
                case "TurnOffAllChannels":
                    return SendMessage(Cmd1_Driver.SET_SWITCH, Cmd2_Driver.SET_SWITCH_ALL_OFF);
                case "Ping":
                    return SendMessage(Cmd1_Driver.PING, Cmd2_Driver.CMD_NULL);
                case "ReadSensors":
                    return SendMessage(Cmd1_Driver.READ, Cmd2_Driver.SENSORS);
                case "ReadBoardPower":
                    return SendMessage(Cmd1_Driver.READ, Cmd2_Driver.BOARD_POWER);
            }
            throw new ArgumentException("Invalid command");
        }
        public static FrameParseResult SendMessage(string message, byte[] state_list)
        {
            switch (message)
            {
                case "SetAllSwitch":
                    return SendMessage(Cmd1_Driver.SET_SWITCH, Cmd2_Driver.SWITCH_ALL, new List<byte>(state_list));
            }
            throw new ArgumentException("Invalid command");
        }
        public static FrameParseResult SendMessage(string message, short channel, bool state)
        {
            switch (message)
            {
                case "SetSwitch":
                    return SendMessage(Cmd1_Driver.SET_SWITCH, (byte)(Cmd2_Driver.CHANNEL_A1 + channel), (byte)(state ? 1 : 0));
            }
            throw new ArgumentException("Invalid command");
        }
        #endregion

        #region Voltage/Current Commands
        public static FrameParseResult SendMessage(string message, short channel, double value)
        {
            switch (message)
            {
                case "SetVoltage":
                    return SendMessage(Cmd1_Driver.SET_VOLTAGE, (byte)(Cmd2_Driver.CHANNEL_A1 + channel), value);
                case "SetCurrent":
                    return SendMessage(Cmd1_Driver.SET_CURRENT, (byte)(Cmd2_Driver.CHANNEL_A1 + channel), value);
            }
            throw new ArgumentException("Invalid command");
        }
        #endregion

        #region Read Status
        public static FrameParseResult SendMessage(string message, short channel)
        {
            switch (message)
            {
                case "ReadChannel":
                    return SendMessage(Cmd1_Driver.READ, (byte)(Cmd2_Driver.CHANNEL_A1 + channel));
                case "GetInitValue":
                    return SendMessage(Cmd1_Driver.GET_INIT_VALUE, (byte)(Cmd2_Driver.CHANNEL_A1 + channel));
            }
            throw new ArgumentException("Invalid command");
        }
        #endregion
    }
}