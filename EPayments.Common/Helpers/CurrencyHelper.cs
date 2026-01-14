using System;
using System.Globalization;

namespace EPayments.Common.Helpers
{
    public enum CurrencyMode
    {
        BGN,
        Dual,
        EUR
    }

    public static class CurrencyHelper
    {
        public static readonly decimal BgnToEuroRate = Decimal.Parse(AppSettings.EPaymentsCommon_EuroExchangeRate, CultureInfo.InvariantCulture);

        public static readonly DateTime EuroAcceptanceDate = AppSettings.EPaymentsCommon_EuroAcceptanceDate;

        public static readonly string CurrencyModeConfig = AppSettings.EpaymentsCommon_CurrencyMode;

        public static bool IsEuroTimePeriod => DateTimeOffset.UtcNow >= EuroAcceptanceDate;

        private static readonly CurrencyMode currencyMode;
        public static CurrencyMode CurrencyMode => currencyMode;

        static CurrencyHelper()
        {
            currencyMode = ParseCurrencyMode(CurrencyModeConfig);
        }

        public static string GetBgnValueFormated(decimal value, DateTime createDate)
        {
            if (createDate >= EuroAcceptanceDate)
            {
                value *= BgnToEuroRate;
            }

            return Formatter.DecimalToTwoDecimalPlacesFormat(value);
        }

        public static decimal GetBgnValue(decimal value, DateTime createDate)
        {
            if (createDate >= EuroAcceptanceDate)
            {
                value *= BgnToEuroRate;
            }

            return value;
        }

        public static string GetEuroValueFormated(decimal value, DateTime createDate)
        {
            if (createDate < EuroAcceptanceDate)
            {
                value /= BgnToEuroRate;
            }

            return Formatter.DecimalToTwoDecimalPlacesFormat(value);
        }

        public static decimal GetEuroValue(decimal value, DateTime createDate)
        {
            if (createDate < EuroAcceptanceDate)
            {
                value /= BgnToEuroRate;
            }

            return value;
        }

        public static string GetCurrencyCode()
        {
            return IsEuroTimePeriod ? "EUR" : "BGN";
        }

        public static string GetCurrencyCodeNumber()
        {
            return IsEuroTimePeriod ? "978" : "975";
        }

        public static string GetValueForDisplayUsingCurrencyMode(decimal value, DateTime createDate)
        {
            if (CurrencyMode == CurrencyMode.Dual)
            {
                if (IsEuroTimePeriod)
                {
                    return $"{GetEuroValueFormated(value, createDate)} € ({GetBgnValueFormated(value, createDate)} лв.)";
                }
                else
                {
                    return $"{GetBgnValueFormated(value, createDate)} лв. ({GetEuroValueFormated(value, createDate)} €)";
                }
            }
            else if (CurrencyMode == CurrencyMode.EUR)
            {
                return $"{GetEuroValueFormated(value, createDate)} €";
            }

            return $"{GetBgnValueFormated(value, createDate)} лв.";
        }

        private static CurrencyMode ParseCurrencyMode(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return CurrencyMode.Dual;

            if (Enum.TryParse<CurrencyMode>(input, ignoreCase: true, out var result))
                return result;

            return CurrencyMode.Dual;
        }

    }
}
