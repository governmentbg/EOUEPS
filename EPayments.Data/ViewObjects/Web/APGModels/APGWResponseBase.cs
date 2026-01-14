using System;
using System.Globalization;

namespace EPayments.Data.ViewObjects.Web.APGModels
{
    public abstract class APGWResponseBase : APGWBaseData
    {
        public abstract bool IsResponseValid();

        protected bool ValidateIntegers(int expected, string actualAsString)
        {
            int actualValue;

            if (!int.TryParse(actualAsString, out actualValue))
            {
                return false;
            }

            if (expected != actualValue)
            {
                return false;
            }

            return true;
        }

        protected bool ValidateDecimal(decimal? expected, string actualAsString)
        {
            decimal actualValue;

            if (expected == null)
            {
                return string.IsNullOrEmpty(actualAsString);
            }

            if (!decimal.TryParse(actualAsString, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out actualValue))
            {
                return false;
            }

            if (expected != actualValue)
            {
                return false;
            }

            return true;
        }

        protected bool ValidateDouble(double expected, string actualAsString)
        {
            double actualValue;

            if (!double.TryParse(actualAsString, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out actualValue))
            {
                return false;
            }

            if (expected != actualValue)
            {
                return false;
            }

            return true;
        }

        protected bool ValidateStrings(string expected, string actual)
        {
            return string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase);
        }

        protected ParsedPart ParseNextPart(string message, int index, int length)
        {
            string newMessage;

            if (int.TryParse(message.Substring(index, length.ToString().Length), out int value))
            {
                int nextIndex = index + length.ToString().Length;

                newMessage = message.Substring(nextIndex, value);

                return new ParsedPart(nextIndex + newMessage.Length, newMessage);
            }
            else
            {
                int nextIndex = index + length.ToString().Length;
                newMessage = message.Substring(index, length);
                if (newMessage == "-")
                    return new ParsedPart(nextIndex, newMessage);
                
                newMessage = message.Substring(nextIndex, length);
                
                if (newMessage == "-")
                {                
                    return new ParsedPart(nextIndex, newMessage);
                }
                else
                    return new ParsedPart(nextIndex + newMessage.Length, newMessage);
            }
        }

        protected internal class ParsedPart
        {
            public ParsedPart(int nextIndex, string message)
            {
                NextIndex = nextIndex;
                Message =  message;
            }

            internal int NextIndex { get; }

            internal string Message { get; }

            public override string ToString()
            {
                return string.Format("Next: {0}, Message: {0}", NextIndex, Message);
            }
        }
    }
}
