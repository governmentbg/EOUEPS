using System;
using System.Text;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.IO;
using System.Globalization;

namespace EPayments.Common.Helpers
{
    public static class Formatter
    {
        public static string ToNotNullString(string value)
        {
            return value != null ? value : String.Empty;
        }

        public static string FormatName(string firstName, string secondName, string lastName)
        {
            return String.Format("{0} {1} {2}", ToNotNullString(firstName), ToNotNullString(secondName), ToNotNullString(lastName)).Trim();
        }

        public static Tuple<string, string, string> SplitNames(string fullName)
        {
            string firstName = String.Empty;
            string secondName = String.Empty;
            string lastName = String.Empty;

            if (!String.IsNullOrWhiteSpace(fullName))
            {
                var splitName = fullName.Split(new Char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                firstName = splitName[0];
                if (splitName.Length > 1)
                {
                    lastName = splitName[splitName.Length - 1];

                    if (splitName.Length > 2)
                    {
                        for (int i = 1; i < splitName.Length - 1; i++)
                        {
                            secondName = secondName + splitName[i] + ' ';
                        }

                        secondName = secondName.Trim();
                    }
                }
            }

            return new Tuple<string, string, string>(firstName, secondName, lastName);
        }

        public static string DateToBgFormat(DateTime dateTime)
        {
            return dateTime.ToString("dd.MM.yyyy");
        }

        public static string DateToBgFormat(DateTime? dateTime)
        {
            if (dateTime.HasValue)
            {
                return DateToBgFormat(dateTime.Value);
            }
            else
            {
                return null;
            }
        }

        public static string DateToBgFormatWithYearSuffix(DateTime dateTime)
        {
            return dateTime.ToString("dd.MM.yyyy г.");
        }

        public static string DateToBgFormatWithYearSuffix(DateTime? dateTime)
        {
            if (dateTime.HasValue)
            {
                return DateToBgFormatWithYearSuffix(dateTime.Value);
            }
            else
            {
                return null;
            }
        }

        public static string DateToBgFormatWithoutYearSuffix(DateTime dateTime)
        {
            return dateTime.ToString("dd.MM.yyyy");
        }

        public static string DateToBgFormatWithoutYearSuffix(DateTime? dateTime)
        {
            if (dateTime.HasValue)
            {
                return DateToBgFormatWithoutYearSuffix(dateTime.Value);
            }
            else
            {
                return null;
            }
        }

        public static string DateTimeToDigitsFormat(DateTime dateTime)
        {
            return dateTime.ToString("yyyyMMddHHmmss");
        }

        public static string DateTimeToBgFormat(DateTime dateTime)
        {
            return dateTime.ToLocalTime().ToString("dd.MM.yyyy г. HH:mm:ss ч.");
        }

        public static string DateTimeToBgFormatNotLocalTime(DateTime dateTime)
        {
            return dateTime.ToString("dd.MM.yyyy г. HH:mm:ss ч.");
        }

        public static string DateTimeToBgFormat(DateTime? dateTime)
        {
            if (dateTime.HasValue)
            {
                return DateTimeToBgFormat(dateTime.Value);
            }
            else
            {
                return null;
            }
        }

        public static string DateTimeToBgFormatWithoutSeconds(DateTime dateTime)
        {
            return dateTime.ToLocalTime().ToString("dd.MM.yyyy, HH:mm");
        }

        public static string DateTimeToBgFormatWithoutSecondsNotLocalTime(DateTime? dateTime)
        {
            if (dateTime.HasValue)
            {
                return ((DateTime)dateTime).ToString("dd.MM.yyyy, HH:mm");
            }
            else
            {
                return null;
            }
        }

        public static string DateTimeToBgFormatWithoutSeconds(DateTime? dateTime)
        {
            if (dateTime.HasValue)
            {
                return DateTimeToBgFormatWithoutSeconds(dateTime.Value.ToLocalTime());
            }
            else
            {
                return null;
            }
        }

        public static string DateTimeToIso8601Format(DateTime dateTime)
        {
            return dateTime.ToString("o");
        }

        public static string DateToShortFormat(DateTime dateTime)
        {
            return dateTime.ToString("yyMMdd");
        }

        public static string DecimalToTwoDecimalPlacesFormat(decimal value)
        {
            return value.ToString("### ### ##0.00", CultureInfo.InvariantCulture).Trim();
        }
        public static string DecimalToTwoDecimalPlacesFormatNoSpaces(decimal value)
        {
            return value.ToString("0.00", CultureInfo.InvariantCulture).Trim();
        }
        public static string DecimalToTwoDecimalPlacesFormat(decimal? value)
        {
            if (value.HasValue)
            {
                return DecimalToTwoDecimalPlacesFormat(value.Value);
            }
            else
            {
                return null;
            }
        }

        public static string EnumToDescriptionString(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());
            DescriptionAttribute[] attributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (attributes.Length > 0)
            {
                return attributes[0].Description;
            }
            else
            {
                return value.ToString();
            }
        }

        public static string TruncateString(string value, int length)
        {
            return value.Length > length ? value.Substring(0, length) : value;
        }

        public static string ConvertTextToHtml(string text)
        {
            //Create a StringBuilder object from the string intput
            //parameter
            StringBuilder sb = new StringBuilder(text);
            //Replace all double white spaces with a single white space
            //and &nbsp;
            sb.Replace("  ", " &nbsp;");
            //Check if HTML tags are not allowed
            //Convert the brackets into HTML equivalents
            sb.Replace("<", "&lt;");
            sb.Replace(">", "&gt;");
            //Convert the double quote
            sb.Replace("\"", "&quot;");
            //Create a StringReader from the processed string of 
            //the StringBuilder object
            StringReader sr = new StringReader(sb.ToString());
            StringWriter sw = new StringWriter();
            //Loop while next character exists
            while (sr.Peek() > -1)
            {
                //Read a line from the string and store it to a temp
                //variable
                string temp = sr.ReadLine();
                //write the string with the HTML break tag
                //Note here write method writes to a Internal StringBuilder
                //object created automatically
                sw.Write(temp + "<br>");
            }
            //Return the final processed text
            return sw.GetStringBuilder().ToString();
        }

        public static string ExceptionToDetailedInfo(Exception ex)
        {
            StringBuilder strBuilder = new StringBuilder();
            if (ex != null)
            {
                GetExceptionInfo(strBuilder, ex);
            }

            return strBuilder.ToString();
        }

        public static Uri UriCombine(string hostUrl, params string[] segments)
        {
            string leadingUri = hostUrl.Trim().Trim(new char[] { '/', '\\' });

            string[] nonEmptySegments = segments.Where(e => !String.IsNullOrWhiteSpace(e)).ToArray();

            if (nonEmptySegments != null && nonEmptySegments.Count() > 0)
            {
                string combinedUri = leadingUri + "/" + nonEmptySegments[0].Trim().Trim(new char[] { '/', '\\' });

                if (nonEmptySegments.Count() == 1)
                {
                    return new Uri(combinedUri);
                }
                else
                {
                    string[] remainingSegments = nonEmptySegments.Skip(1).ToArray();

                    return UriCombine(combinedUri, remainingSegments);
                }
            }

            return new Uri(leadingUri);
        }

        public static string UriSegmentsCombine(string segment, params string[] segments)
        {
            string leadingUri = segment.Trim().Trim(new char[] { '/', '\\' });

            string[] nonEmptySegments = segments.Where(e => !String.IsNullOrWhiteSpace(e)).ToArray();

            if (nonEmptySegments != null && nonEmptySegments.Count() > 0)
            {
                string combinedUri = leadingUri + "/" + nonEmptySegments[0].Trim().Trim(new char[] { '/', '\\' });

                if (nonEmptySegments.Count() == 1)
                {
                    return combinedUri;
                }
                else
                {
                    string[] remainingSegments = nonEmptySegments.Skip(1).ToArray();

                    return UriSegmentsCombine(combinedUri, remainingSegments);
                }
            }

            return leadingUri;
        }

        public static string AddLeadingSymbols(string value, int length, char symbol)
        {
            int valueLength = value.Length;

            if (valueLength >= length)
                return value;

            for (int i = 0; i < (length - valueLength); i++)
            {
                value = string.Format("{0}{1}", symbol.ToString(), value);
            }

            return value;
        }

        public static string AddEndingSymbols(string value, int length, char symbol)
        {
            int valueLength = value.Length;

            if (valueLength >= length)
                return value;

            for (int i = 0; i < (length - valueLength); i++)
            {
                value = string.Format("{0}{1}", value, symbol.ToString());
            }

            return value;
        }

        public static string CamelCaseEachWord(string value)
        {
            if (!String.IsNullOrWhiteSpace(value))
            {
                TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                return textInfo.ToTitleCase(value.ToLower());
            }
            else
            {
                return String.Empty;
            }
        }

        public static string WhiteSpaceToNull(string value)
        {
            return !String.IsNullOrWhiteSpace(value) ? value : null;
        }

        private static void GetExceptionInfo(StringBuilder stringBuilder, Exception exception)
        {
            stringBuilder.AppendFormat("Exception type: {0}\n", exception.GetType().FullName);
            stringBuilder.AppendFormat("Message: {0}\n", exception.Message);
            stringBuilder.AppendFormat("Stack trace:\n{0}\n", exception.StackTrace);
            if (exception.InnerException != null)
            {
                stringBuilder.AppendFormat("\n\nInner Exception:\n");
                GetExceptionInfo(stringBuilder, exception.InnerException);
            }
        }

    }
}