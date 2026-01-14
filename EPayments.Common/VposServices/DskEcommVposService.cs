using EPayments.Common.Helpers;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Web;

namespace EPayments.Common.VposHelpers
{
    public enum DskEcommVposTransactionResult
    {
        Pending = 0,
        Blocked = 1,
        Successful = 2,
        AuthorizationCancelled = 3,
        TransactionRestored = 4,
        InitializedAuthorizationByACS = 5,
        AuthorizationRejected = 6

        //0 Поръчката е регистрирана, но все още не е платена.
        //1 Блокирана е сума преди оторизацията (за двуфазно плащане)
        //2 Сумата e депозирана успешно
        //3 Оторизацията е отменена
        //4 Транзакцията е възстановена
        //5 Оторизацията чрез ACS на издадетеля е инициирана
        //6 Оторизацията е отказана
    }

    public class DskEcommVposRegisteredOrder
    {
        public string OrderId { get; set; }
        public string FormUrl { get; set; }
        public string ErrorCode { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class DskEcommVposOrderStatus
    {
        public string ApprovalCode { get; set; }
        public int? AuthCode { get; set; }
        public string OrderNumber { get; set; }
        public int? Amount { get; set; }
        public string ErrorCode { get; set; }
        public string ErrorMessage { get; set; }

        public DskEcommVposTransactionResult? OrderStatus { get; set; }
        
    }

    public static class DskEcommVposService
    {
        public static DskEcommVposRegisteredOrder RegisterOrder(
            string vposEcommBaseUrl,
            string merchantUsername,
            string merchantPassword,
            string requestIdentifier,
            decimal paymentAmount,
            string paymentReason)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(vposEcommBaseUrl);

                List<KeyValuePair<string, string>> bodyParams = new List<KeyValuePair<string, string>>();

                bodyParams.Add(new KeyValuePair<string, string>("userName", merchantUsername));
                bodyParams.Add(new KeyValuePair<string, string>("password", merchantPassword));
                bodyParams.Add(new KeyValuePair<string, string>("orderNumber", $"{requestIdentifier}_{Guid.NewGuid().ToString().Substring(0, 8)}"));
                bodyParams.Add(new KeyValuePair<string, string>("amount", AppSettings.EPaymentsWeb_PaymentUseMinTestAmount ? "50" : ((int)Math.Truncate(100 * paymentAmount)).ToString()));
                bodyParams.Add(new KeyValuePair<string, string>("currency", CurrencyHelper.GetCurrencyCodeNumber()));
                bodyParams.Add(new KeyValuePair<string, string>("returnUrl",
                    Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, "Vpos/DskEcommPaymentResult").ToString()));
                bodyParams.Add(new KeyValuePair<string, string>("failUrl",
                    Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, "Vpos/DskEcommPaymentResult").ToString()));
                bodyParams.Add(new KeyValuePair<string, string>("description", paymentReason.Length > 256 ? paymentReason.Substring(0, 256) : paymentReason));

                string encodedQueryString = String.Join("&", bodyParams.Select(e => $"{HttpUtility.UrlEncode(e.Key)}={HttpUtility.UrlEncode(e.Value)}"));

                DisableServerCertificateValidationAndSetTlsVersion();

                var response = client.GetAsync($"payment/rest/register.do?{encodedQueryString}").Result;

                if (response.IsSuccessStatusCode)
                {
                    var reponseContent = response.Content.ReadAsStringAsync().Result;

                    return JsonConvert.DeserializeObject<DskEcommVposRegisteredOrder>(reponseContent);
                }
                else
                {
                    throw new Exception("Error occured when try to register order. HTTPStatusCode: " + response.StatusCode.ToString());
                }
            }
        }

        public static Tuple<DskEcommVposOrderStatus, string> CheckOrderStatus(
            string vposEcommBaseUrl,
            string merchantUsername,
            string merchantPassword,
            string orderId)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(vposEcommBaseUrl);

                List<KeyValuePair<string, string>> bodyParams = new List<KeyValuePair<string, string>>();

                bodyParams.Add(new KeyValuePair<string, string>("userName", merchantUsername));
                bodyParams.Add(new KeyValuePair<string, string>("password", merchantPassword));
                bodyParams.Add(new KeyValuePair<string, string>("orderId", orderId));

                string encodedQueryString = String.Join("&", bodyParams.Select(e => $"{HttpUtility.UrlEncode(e.Key)}={HttpUtility.UrlEncode(e.Value)}"));

                DisableServerCertificateValidationAndSetTlsVersion();

                var response = client.GetAsync($"payment/rest/getOrderStatus.do?{encodedQueryString}").Result;

                if (response.IsSuccessStatusCode)
                {
                    var reponseContent = response.Content.ReadAsStringAsync().Result;

                    DskEcommVposOrderStatus orderStatus = JsonConvert.DeserializeObject<DskEcommVposOrderStatus>(reponseContent);

                    return new Tuple<DskEcommVposOrderStatus, string>(orderStatus, reponseContent);
                }
                else
                {
                    throw new Exception("Error occured when try to check status. HTTPStatusCode: " + response.StatusCode.ToString());
                }
            }
        }

        private static void DisableServerCertificateValidationAndSetTlsVersion()
        {
            ServicePointManager.ServerCertificateValidationCallback +=
                delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

            System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        }
    }
}
