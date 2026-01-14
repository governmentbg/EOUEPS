using System;
using System.Text;

namespace EPayments.Common.Helpers
{
    public static class BinHexHelper
    {
        public static string ByteArrayToHexString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }

        public static byte[] HexStringToByteArray(string hexString)
        {
            byte[] signature = new byte[hexString.Length / 2];

            for (int i = 0; i < hexString.Length; i += 2)
            {
                int value = HexToBin(hexString[i]) * 16 + HexToBin(hexString[i + 1]);
                signature[i / 2] = (byte)value;
            }

            return signature;
        }

        private static byte HexToBin(char next)
        {
            byte value;

            if (char.IsLetter(next))
            {
                if (char.IsUpper(next))
                {
                    value = (byte)(next - 'A' + 10);
                }
                else
                {
                    value = (byte)(next - 'a' + 10);
                }
            }
            else if (!char.IsDigit(next))
            {
                throw new Exception();
            }
            else
            {
                value = (byte)(next - '0');
            }

            if (value < 0 || 15 < value)
            {
                throw new InvalidOperationException("Provided char is not Hex number");
            }

            return value;
        }
    }
}
