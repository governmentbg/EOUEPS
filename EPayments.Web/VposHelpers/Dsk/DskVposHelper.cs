using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Models;
using EPayments.Web.Common;
using EPayments.Web.Controllers;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.VposHelpers.Dsk
{
    public static class DskVposHelper
    {
        public static AutoPostVM CreateAutoPostVM(VposRequestDataVO dataVO, Guid? vposRedirectGid)
        {
            var bodyParams = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("merchant_id", dataVO.DskVposMerchantId),
                new KeyValuePair<string, string>("amount", AppSettings.EPaymentsWeb_PaymentUseMinTestAmount ? "0.01" : Formatter.DecimalToTwoDecimalPlacesFormat(dataVO.PaymentAmount)),
                new KeyValuePair<string, string>("currency_code", CurrencyHelper.GetCurrencyCode()),
                new KeyValuePair<string, string>("track_id", String.Format("{0}_{1}", dataVO.PaymentRequestIdentifier.ToString(), Guid.NewGuid().ToString().Substring(0, 8))),
                new KeyValuePair<string, string>("description", dataVO.PaymentReason.Length > 256 ? dataVO.PaymentReason.Substring(0, 256) : dataVO.PaymentReason),
                new KeyValuePair<string, string>("payment_timeout", "10"),
                new KeyValuePair<string, string>("ok_return_url", 
                    Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Vpos.Name, MVC.Vpos.ActionNames.DskSuccessfulPayment).ToString() +
                    String.Format("?gid={0}&vposRedirectGid={1}", dataVO.ЕserviceClientGid.ToString().ToLower(), vposRedirectGid.HasValue ? vposRedirectGid.ToString().ToLower() : String.Empty)),
                new KeyValuePair<string, string>("cancel_return_url", 
                    Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Vpos.Name, MVC.Vpos.ActionNames.DskFailedPayment).ToString() +
                    String.Format("?gid={0}&vposRedirectGid={1}", dataVO.ЕserviceClientGid.ToString().ToLower(), vposRedirectGid.HasValue ? vposRedirectGid.ToString().ToLower() : String.Empty)),
                new KeyValuePair<string, string>("language", "bg-BG"),
                new KeyValuePair<string, string>("merchant_timestamp", DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss")),
                new KeyValuePair<string, string>("is_unique", AppSettings.EPaymentsWeb_PaymentAllowDublicateTrackId ? "0" : "1"),
            };

            //calculate header value
            string paramsNames = "";
            string paramsLengths = "";
            int paramsCount = 0;
            for (int i = 0; i < bodyParams.Count; i++)
            {
                paramsNames += bodyParams[i].Key + ",";
                paramsLengths += bodyParams[i].Value.Length.ToString("D3");
                paramsCount++;
            }
            var header = String.Format("{0}{1}{2}", paramsCount.ToString("D2"), paramsNames, paramsLengths);

            //calculate inputString
            string inputString = "";
            inputString += header;
            for (int i = 0; i < bodyParams.Count; i++) //adding params value
            {
                inputString += bodyParams[i].Value;
            }
            inputString += dataVO.DskVposMerchantPassword; //adding password

            //genarate mac code
            var mac = CryptoHelper.CalculateMD5Hash(inputString);

            bodyParams.Add(new KeyValuePair<string, string>("header", header));
            bodyParams.Add(new KeyValuePair<string, string>("mac", mac));

            AutoPostVM model = new AutoPostVM();
            model.PostUrl = dataVO.VposPaymentRequestUri;
            model.PostValues = bodyParams;

            return model;
        }

        public static bool IsRequestValid(EserviceClient еserviceClient, DskVposResultDO resultDO)
        {
            if (AppSettings.EPaymentsWeb_SimulateVposPayment)
            {
                return true;
            }

            try
            {
                if (еserviceClient == null || String.IsNullOrWhiteSpace(resultDO.Return_header) || String.IsNullOrWhiteSpace(resultDO.Return_mac))
                {
                    return false;
                }

                var splittedValues = resultDO.Return_header.Substring(2).Split(new char[] { ',' });

                var paramsCountStr = resultDO.Return_header.Substring(0, 2);
                var paramsLengthStr = splittedValues.Last();

                List<string> paramNames = new List<string>();
                for (int i = 0; i < splittedValues.Length - 1; i++)
                {
                    paramNames.Add(splittedValues[i]);
                }

                if (int.Parse(paramsCountStr) != paramNames.Count)
                {
                    return false;
                }

                foreach (var name in paramNames)
                {
                    if (!IsParameterNameValid(name))
                    {
                        return false;
                    }
                }

                for (int i = 0; i < paramNames.Count; i++)
                {
                    var paramValue = GetParameterValue(paramNames[i], resultDO);

                    var expectedParamLength = int.Parse(paramsLengthStr.Substring(i * 3, 3));

                    if (paramValue.Length != expectedParamLength)
                    {
                        return false;
                    }
                }

                string inputString = "";
                inputString += resultDO.Return_header;
                for (int i = 0; i < paramNames.Count; i++)
                {
                    var paramValue = GetParameterValue(paramNames[i], resultDO);

                    inputString += paramValue;
                }
                inputString += еserviceClient.DskVposMerchantPassword;

                var expectedMac = CryptoHelper.CalculateMD5Hash(inputString);

                if (expectedMac.ToLower() != resultDO.Return_mac.ToLower())
                {
                    return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsParameterNameValid(string paramName)
        {
            var name = paramName.ToLower();

            return
                name == "payment_id" ||
                name == "track_id" ||
                name == "auth_code" ||
                name == "reference_id" ||
                name == "transaction_id" ||
                name == "post_date" ||
                name == "result" ||
                name == "response_code" ||
                name == "error_code" ||
                name == "status";
        }

        private static string GetParameterValue(string paramName, DskVposResultDO resultDO)
        {
            var name = paramName.ToLower();

            switch (name)
            {
                case "payment_id":
                    return resultDO.Payment_id ?? String.Empty;
                case "track_id":
                    return resultDO.Track_id ?? String.Empty;
                case "auth_code":
                    return resultDO.Auth_code ?? String.Empty;
                case "reference_id":
                    return resultDO.Reference_id ?? String.Empty;
                case "transaction_id":
                    return resultDO.Transaction_id ?? String.Empty;
                case "post_date":
                    return resultDO.Post_date ?? String.Empty;
                case "result":
                    return resultDO.Result ?? String.Empty;
                case "response_code":
                    return resultDO.Response_code ?? String.Empty;
                case "error_code":
                    return resultDO.Error_code ?? String.Empty;
                case "status":
                    return resultDO.Status ?? String.Empty;
                default:
                    throw new ArgumentException();
            }
        }
    }
}