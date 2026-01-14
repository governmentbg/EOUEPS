using System;
using System.Text;
using System.Linq;
using System.Reflection;
using System.ComponentModel;
using System.IO;
using System.Globalization;

namespace EPayments.Common.Helpers
{
    public enum DateFormat
    {
        yy_MM_dd_NoSeparator,
    }

    public static class Parser
    {
        public static DateTime? BgFormatDateStringToDateTime(string dateValue)
        {
            DateTime? returnValue = null;

            if (!String.IsNullOrWhiteSpace(dateValue))
            {
                returnValue = DateTime.ParseExact(dateValue.Trim(), "dd.MM.yyyy", CultureInfo.InvariantCulture).Date;
            }

            return returnValue;
        }

        public static decimal? TwoDecimalPlacesFormatStringToDecimal(string decimalValue)
        {
            decimal? returnValue = null;

            if (!String.IsNullOrWhiteSpace(decimalValue))
            {
                returnValue = decimal.Parse(decimalValue, CultureInfo.InvariantCulture);
            }

            return returnValue;
        }

        public static DateTime? GetDateFirstMinute(DateTime? dateValue)
        {
            DateTime? returnValue = null;

            if (dateValue.HasValue)
            {
                returnValue = dateValue.Value.Date;
            }

            return returnValue;
        }

        public static DateTime? GetDateLastMinute(DateTime? dateValue)
        {
            DateTime? returnValue = null;

            if (dateValue.HasValue)
            {
                returnValue = dateValue.Value.Date.AddDays(1).AddTicks(-1);
            }

            return returnValue;
        }

        public static DateTime? DateParse(string content, DateFormat format)
        {
            try
            {
                switch (format)
                {
                    case DateFormat.yy_MM_dd_NoSeparator:
                        {
                            return DateTime.ParseExact(content, "yyMMdd", CultureInfo.InvariantCulture);
                        }
                    default:
                        throw new ArgumentException("invalid format");
                }
            }
            catch
            {
                return null;
            }
        }
    }
}