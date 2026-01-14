using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Model.Models;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Text.RegularExpressions;

namespace EPayments.Web.VposHelpers.Epay
{
    public class EpayVposHelper
    {
        public static AutoPostVM CreateAutoPostVM(string invoiceNo, PaymentRequest paymentRequest)
        {
            string pattern = @"[^а-яА-Яa-zA-Z0-9\s\-,\.]";

            List<KeyValuePair<string, string>> postData = new List<KeyValuePair<string, string>>();
            bool isMultilineIBAN = false;
            if (paymentRequest.ServiceProviderIBAN.Length >= 12) 
            {
                isMultilineIBAN = paymentRequest.ServiceProviderIBAN[12] == '8';
            }

            postData.Add(new KeyValuePair<string, string>("MIN", AppSettings.EPaymentsWeb_EpayVposKin));
            postData.Add(new KeyValuePair<string, string>("INVOICE", invoiceNo));
            postData.Add(new KeyValuePair<string, string>("MERCHANT", ReplaceInvalidCharacters(paymentRequest.ServiceProviderName, pattern)));
            postData.Add(new KeyValuePair<string, string>("IBAN", paymentRequest.ServiceProviderIBAN));
            postData.Add(new KeyValuePair<string, string>("BIC", paymentRequest.ServiceProviderBIC));
            if (CurrencyHelper.IsEuroTimePeriod)
            {
                postData.Add(new KeyValuePair<string, string>("CURRENCY", "EUR"));

            }
            if (isMultilineIBAN)
            {
                postData.Add(new KeyValuePair<string, string>("SUM1", Formatter.DecimalToTwoDecimalPlacesFormat(paymentRequest.PaymentAmount)));
                postData.Add(new KeyValuePair<string, string>("TOTAL", Formatter.DecimalToTwoDecimalPlacesFormat(paymentRequest.PaymentAmount)));
                postData.Add(new KeyValuePair<string, string>("PSTATEMENT1", "448007"));
                postData.Add(new KeyValuePair<string, string>("STATEMENT1", TrimStringLength(ReplaceInvalidCharacters(paymentRequest.PaymentReason, pattern), 52)));
                postData.Add(new KeyValuePair<string, string>("DOC_NO1", "9"));
            }
            else
            {
                postData.Add(new KeyValuePair<string, string>("AMOUNT", Formatter.DecimalToTwoDecimalPlacesFormat(paymentRequest.PaymentAmount)));
                postData.Add(new KeyValuePair<string, string>("STATEMENT", TrimStringLength(ReplaceInvalidCharacters(paymentRequest.PaymentReason, pattern), 52)));
                postData.Add(new KeyValuePair<string, string>("DOC_NO", "9"));
            }

            postData.Add(new KeyValuePair<string, string>("EXP_TIME", Formatter.DateToBgFormat(DateTime.Now.AddHours(1))));
            postData.Add(new KeyValuePair<string, string>("DESCR", paymentRequest.PaymentRequestIdentifier));
            postData.Add(new KeyValuePair<string, string>("OBLIG_PERSON", TrimStringLength(ReplaceInvalidCharacters(paymentRequest.ApplicantName, pattern), 26)));

            if (paymentRequest.ApplicantUinTypeId == Model.Enums.UinType.Egn)
            {
                postData.Add(new KeyValuePair<string, string>("EGN", paymentRequest.ApplicantUin));
            }
            else if (paymentRequest.ApplicantUinTypeId == Model.Enums.UinType.Bulstat)
            {
                postData.Add(new KeyValuePair<string, string>("BULSTAT", paymentRequest.ApplicantUin));
            }
            else if (paymentRequest.ApplicantUinTypeId == Model.Enums.UinType.Lnch)
            {
                postData.Add(new KeyValuePair<string, string>("LNC", paymentRequest.ApplicantUin));
            }

            string data = String.Join(Environment.NewLine, postData.Select(e => $"{e.Key}={e.Value}"));
            var dataBytes = Encoding.GetEncoding("windows-1251").GetBytes(data);
            string encoded = Convert.ToBase64String(dataBytes);
            string checksum = CalculateHmac(AppSettings.EPaymentsWeb_EpayVposSecret, encoded);

            AutoPostVM model = new AutoPostVM();
            model.PostUrl = AppSettings.EPaymentsWeb_EpayVposUrl;
            model.PostValues = new List<KeyValuePair<string, string>>();
            model.PostValues.Add(new KeyValuePair<string, string>("PAGE", "paylogin"));
            model.PostValues.Add(new KeyValuePair<string, string>("ENCODED", encoded));
            model.PostValues.Add(new KeyValuePair<string, string>("CHECKSUM", checksum));
            model.PostValues.Add(new KeyValuePair<string, string>("URL_OK", AppSettings.EPaymentsWeb_EpayOK + "&ino=" + invoiceNo));
            model.PostValues.Add(new KeyValuePair<string, string>("URL_CANCEL", AppSettings.EPaymentsWeb_EpayCancel + "&ino=" + invoiceNo));

            return model;
        }

        public static string CalculateHmac(string secretCode, string content)
        {
            string hmac;

            var secretCodeBytes = Encoding.UTF8.GetBytes(secretCode);
            var contentBytes = Encoding.UTF8.GetBytes(content);

            using (var hmacSHA1 = new HMACSHA1(secretCodeBytes))
            {
                var hash = hmacSHA1.ComputeHash(contentBytes);

                hmac = BitConverter.ToString(hash).Replace("-", "").ToLower();
            }

            return hmac;
        }

        public static EpayResultData GetEpayResultData(string resultData)
        {
            List<string> values = resultData.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            string invoiceStr = values.Single(e => e.TrimStart().ToLower().StartsWith("invoice=")).Trim().Substring("invoice=".Length);
            string statusStr = values.Single(e => e.TrimStart().ToLower().StartsWith("status=")).Trim().Substring("status=".Length);

            EpayPaymentStatus status = (EpayPaymentStatus)Enum.Parse(typeof(EpayPaymentStatus), statusStr, true);

            return new EpayResultData()
            {
                Invoice = invoiceStr,
                Status = status
            };
        }

        private static string TrimStringLength(string content, int maxSymbols)
        {
            if (content.Length > maxSymbols)
            {
                return content.Substring(0, maxSymbols);
            }
            else
            {
                return content;
            }
        }

        private static string ReplaceInvalidCharacters(string input, string pattern, string replacement = "")
        {
            string result = Regex.Replace(input, pattern, replacement);

            return result;
        }
    }
}