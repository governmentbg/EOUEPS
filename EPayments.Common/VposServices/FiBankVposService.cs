using EPayments.Common.Helpers;
using java.util;
using lv.tietoenator.cs.ecomm.merchant;
using System;
using System.IO;

namespace EPayments.Common.VposHelpers
{
    public enum FiBankVposTransactionResult
    {
        Pending = 1,
        Unsuccessful = 2,
        Successful = 3,
    }

    public static class FiBankVposService
    {
        public static string GenerateTransactionId(
            string fiBankVposMerchantHandlerUrl,
            string clientIpAddress,
            string fiBankVposAccessKeystoreFilename,
            string fiBankVposAccessKeystorePassword,
            decimal paymentAmount,
            string description)
        {
            string paymentAmountFormattedStr = ((int)Math.Truncate(100 * paymentAmount)).ToString();
            string bgCurrencyCode = CurrencyHelper.GetCurrencyCodeNumber();

            Merchant merchantClient = CreateMerchantClient(
                fiBankVposAccessKeystoreFilename,
                fiBankVposAccessKeystorePassword,
                fiBankVposMerchantHandlerUrl);

            string transactionResult = merchantClient.startSMSTrans(
                paymentAmountFormattedStr,
                bgCurrencyCode,
                clientIpAddress,
                description);

            string transactionId = null;

            if (!String.IsNullOrWhiteSpace(transactionResult) && transactionResult.Trim().ToUpper().StartsWith("TRANSACTION_ID:"))
            {
                transactionId = transactionResult.Substring("TRANSACTION_ID:".Length).Trim();
            }

            if (String.IsNullOrWhiteSpace(transactionId))
            {
                throw new Exception("Failed generating Fibank transaction id. Result: " + transactionResult);
            }

            return transactionId;
        }

        public static Tuple<FiBankVposTransactionResult, string> CheckTransactionResult(
            string fiBankVposMerchantHandlerUrl,
            string clientIpAddress,
            string fiBankVposAccessKeystoreFilename,
            string fiBankVposAccessKeystorePassword,
            string transactionId)
        {
            FiBankVposTransactionResult? result = null;
            string resultData = null;

            Merchant merchantClient = CreateMerchantClient(
                fiBankVposAccessKeystoreFilename,
                fiBankVposAccessKeystorePassword,
                fiBankVposMerchantHandlerUrl);

            resultData = merchantClient.getTransResult(transactionId, clientIpAddress);

            var rows = resultData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            if (rows[0].ToLower().StartsWith("result:"))
            {
                string resultStr = rows[0].Substring("result:".Length).ToLower().Trim();

                if (!String.IsNullOrWhiteSpace(resultStr))
                {
                    switch (resultStr)
                    {
                        case "created"://плащане, регистрирано в банковата система (незавършено)
                        case "pending"://плащане, чакащо да бъде платено (незавършено)
                            {
                                result = FiBankVposTransactionResult.Pending;
                                break;
                            }
                        case "timeout"://времето за плащане е изтекло (незавършено)
                        case "failed"://неуспешно плащане
                        case "declined"://отказано плащане от ECOMM
                        case "reversed"://анулирано (отменено) плащане
                        case "autoreversed"://плащане, отменено чрез автоматична анулация
                            {
                                result = FiBankVposTransactionResult.Unsuccessful;
                                break;
                            }
                        case "ok"://успешно плащане
                            {
                                result = FiBankVposTransactionResult.Successful;
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
                //isApproved = result == "ok";
            }

            if (!result.HasValue)
                throw new Exception("Invalid transaction result data");

            return new Tuple<FiBankVposTransactionResult, string>(result.Value, resultData);
        }

        private static Merchant CreateMerchantClient(
            string fiBankVposAccessKeystoreFilename,
            string fiBankVposAccessKeystorePassword,
            string fiBankVposMerchantHandlerUrl)
        {
            var merchantProperties = new Properties();

            string certificatePath = Path.Combine(AppSettings.EPaymentsCommon_FiBankCertificateFolder, fiBankVposAccessKeystoreFilename);

            merchantProperties.setProperty("bank.server.url", fiBankVposMerchantHandlerUrl);
            merchantProperties.setProperty("keystore.file", certificatePath);
            merchantProperties.setProperty("keystore.type", "JKS");
            merchantProperties.setProperty("keystore.password", fiBankVposAccessKeystorePassword);

            return new Merchant(merchantProperties);
        }
    }
}
