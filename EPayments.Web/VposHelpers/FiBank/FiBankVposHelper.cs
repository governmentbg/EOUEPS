using EPayments.Common;
using EPayments.Common.VposHelpers;
using EPayments.Data.ViewObjects.Web;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.VposHelpers.FiBank
{
    public static class FiBankVposHelper
    {
        public static string GenerateTransactionId(VposRequestDataVO dataVO, string clientIpAddress)
        {
            decimal paymentAmount = AppSettings.EPaymentsWeb_PaymentUseMinTestAmount ? 1.00M : dataVO.PaymentAmount;
            string paymentDescription = $"{dataVO.PaymentRequestIdentifier}_{DateTimeOffset.Now.ToUnixTimeSeconds()}";

            return FiBankVposService.GenerateTransactionId(
                AppSettings.EPaymentsWeb_FiBankVposMerchantHandlerUrl,
                clientIpAddress,
                dataVO.FiBankVposAccessKeystoreFilename,
                dataVO.FiBankVposAccessKeystorePassword,
                paymentAmount,
                paymentDescription);
        }

        public static Tuple<FiBankVposTransactionResult, string> CheckTransactionResult(
            string fiBankVposAccessKeystoreFilename,
            string fiBankVposAccessKeystorePassword,
            string transactionId,
            string clientIpAddress)
        {
            return FiBankVposService.CheckTransactionResult(
                AppSettings.EPaymentsWeb_FiBankVposMerchantHandlerUrl,
                clientIpAddress,
                fiBankVposAccessKeystoreFilename,
                fiBankVposAccessKeystorePassword,
                transactionId);
        }

        public static AutoPostVM CreateAutoPostVM(string transactionId)
        {
            var bodyParams = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("trans_id", transactionId)
            };

            AutoPostVM model = new AutoPostVM();
            model.PostUrl = AppSettings.EPaymentsWeb_FiBankVposClientHandlerUrl;
            model.PostValues = bodyParams;

            return model;
        }
    }
}