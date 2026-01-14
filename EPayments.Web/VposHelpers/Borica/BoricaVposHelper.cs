using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Data.ViewObjects.Web;
using EPayments.Data.ViewObjects.Web.APGModels;
using EPayments.Data.ViewObjects.Web.APGModels.Requests;
using EPayments.Model.Models;
using EPayments.Web.Common;
using EPayments.Web.Controllers;
using EPayments.Web.Models.Shared;
using EPayments.Web.VposHelpers.Borica;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace EPayments.Web.VposHelpers.Borica
{
    public static class BoricaVposHelper
    {
        private static Encoding EncodingWindows1251 = Encoding.GetEncoding("windows-1251");

        public static string CreateBOReqRedirectUrl(VposRequestDataVO dataVO, string boricaRequestIdentifier)
        {
            StringBuilder requestData = new StringBuilder();

            //TransactionType
            requestData.Append("10");
            //DateTime
            requestData.Append(Formatter.DateTimeToDigitsFormat(DateTime.Now));
            //Amount
            requestData.Append(AppSettings.EPaymentsWeb_PaymentUseMinTestAmount ? "000000000001" :
                Formatter.AddLeadingSymbols(Regex.Replace(Formatter.DecimalToTwoDecimalPlacesFormat(dataVO.PaymentAmount).Replace(".", String.Empty), @"\s+", ""), 12, '0'));
            //TerminalId
            requestData.Append(Formatter.AddEndingSymbols(dataVO.BoricaVposTerminalId, 8, ' '));
            //PaymentIdentifier
            requestData.Append(Formatter.AddEndingSymbols(boricaRequestIdentifier, 15, ' '));
            //Description
            requestData.Append(Formatter.AddEndingSymbols(String.Format("№{0} {1}", dataVO.PaymentRequestIdentifier, dataVO.PaymentReason ?? String.Empty), 125, ' ').Substring(0, 125));
            //Language
            requestData.Append("BG");
            //ProtocolVersion
            requestData.Append("1.0");

            //Add sign data
            string signData = SignBOReqRequestData(requestData.ToString(), dataVO.BoricaVposBOReqSignCertFileName, dataVO.BoricaVposBOReqSignCertPassword);
            requestData.Append(signData);

            string requestDataBase64 = Convert.ToBase64String(EncodingWindows1251.GetBytes(requestData.ToString()));

            string requestDataBase64Encoded = HttpUtility.UrlEncode(requestDataBase64);

            string redirectUrl = String.Format("{0}?eBorica={1}", dataVO.VposPaymentRequestUri, requestDataBase64Encoded);

            return redirectUrl;
        }

        public static APGWPaymentRequestDataVO CreateBoricaRequestModel(
            VposRequestDataVO dataVO, 
            int boricaRequestIdentifier,
            EserviceClient eserviceClient,
            X509Certificate2 boricaCertificate)
        {
            string boricaRequestIdentifierAsString = boricaRequestIdentifier.ToString("000000");
            string reason;
            string objectName = "O";
            string CNName = "CN";
            string emailName = "E";

            if (dataVO.PaymentReason == null)
            {
                reason = AppSettings.EPaymentsWeb_CentralVposDescription;
            }
            else
            {
                reason = dataVO.PaymentReason.Length <= 50 ? dataVO.PaymentReason : dataVO.PaymentReason.Substring(0, 50);
            }

            Dictionary<string, string> keyValues = GetCertificateParameters(boricaCertificate);

            APGWPaymentRequestDataVO apgwRequest = new APGWPaymentRequestDataVO()
            {
                Merchant_Name = keyValues[objectName],
                Merchant = eserviceClient.BoricaVposMerchantId,
                DESC = reason,
                Amount = AppSettings.EPaymentsWeb_PaymentUseMinTestAmount ? 0.01M : dataVO.PaymentAmount,
                Merchant_Url = keyValues[CNName],
                Order = boricaRequestIdentifierAsString,
                AdCustomBorOrderId = boricaRequestIdentifier.ToString("000000") + AppSettings.EPaymentsWeb_CentralVposPrefixHelper + boricaRequestIdentifier.ToString("00000000000000"),
                ADDENDUM = AppSettings.EPaymentsWeb_CentralVposADDENDUM,
                Date = DateTime.Now.ToUniversalTime(),
                Terminal = eserviceClient.BoricaVposTerminalId,
                Email = keyValues[emailName]
            };

            return apgwRequest;
        }

        public static BoricaVposResultDO GetBoricaVposResultDOFromBOResp(string eBorica)
        {
            string eBoricaDecoded = EncodingWindows1251.GetString(Convert.FromBase64String(eBorica));

            BoricaVposResultDO resultDO = new BoricaVposResultDO(eBoricaDecoded);

            return resultDO;
        }

        public static BoricaVposResultDO GetBoricaVposResultDOFromBOResp(APGWPaymentResponseDataDO apgwResponse)
        {
            return new BoricaVposResultDO(apgwResponse);
        }

        public static bool IsBORespRequestValid(string eBorica, string certificateFileName)
        {
            bool isValid = false;

            try
            {
                if (!String.IsNullOrWhiteSpace(eBorica))
                {
                    string eBoricaDecoded = EncodingWindows1251.GetString(Convert.FromBase64String(eBorica));
                    if (eBoricaDecoded.Length > 56)
                    {
                        string originalMessage = eBoricaDecoded.Substring(0, 56);
                        string signedMessage = eBoricaDecoded.Substring(56);

                        string certificatePath = Path.Combine(AppSettings.EPaymentsWeb_BoricaCertificateFolder, certificateFileName);

                        byte[] bytes = File.ReadAllBytes(certificatePath);
                        X509Certificate2 cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(bytes);

                        using (RSACryptoServiceProvider rsaProvider = (RSACryptoServiceProvider)cert.PublicKey.Key)
                        {
                            byte[] bytesToVerify = EncodingWindows1251.GetBytes(originalMessage);

                            byte[] signatureBytes = EncodingWindows1251.GetBytes(signedMessage);

                            isValid = rsaProvider.VerifyData(bytesToVerify, "SHA1", signatureBytes);
                        }
                    }
                }
            }
            catch
            {
            }

            return isValid;
        }

        public static bool IsBORespRequestValid(APGWPaymentResponseDataDO apgwResponse, string certificateFileName)
        {
            bool isValid = false;

            if (apgwResponse.IsResponseValid())
            {
                byte[] signature = BinHexHelper.HexStringToByteArray(apgwResponse.P_Sign);

                var data = Encoding.UTF8.GetBytes(apgwResponse.GetPSignData());

                string certificatePath = Path.Combine(AppSettings.EPaymentsWeb_BoricaCertificateFolder, certificateFileName);

                X509Certificate2Collection collection = new X509Certificate2Collection();
                try
                {
                    collection.Import(Path.Combine(certificatePath));

                    using (X509Certificate2 certificate = collection[0])
                    {
                        using (var rsa = certificate.GetRSAPublicKey())
                        {
                            isValid = rsa.VerifyData(data, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                        }
                    }
                }
                catch
                {
                    
                }
            }

            return isValid;
        }

        private static string SignBOReqRequestData(string requestData, string certificateFileName, string certificatePassword)
        {
            X509Certificate2 cert;
            try
            {
                string certificatePath = Path.Combine(AppSettings.EPaymentsWeb_BoricaCertificateFolder, certificateFileName);

                byte[] bytes = File.ReadAllBytes(certificatePath);
                cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(bytes, certificatePassword,
                    AppSettings.EPaymentsCommon_UseMachineKeySet == false ?
                    X509KeyStorageFlags.PersistKeySet :
                    X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.MachineKeySet);

                using (RSACryptoServiceProvider rsaProvider = (RSACryptoServiceProvider)cert.PrivateKey)
                {
                    if (rsaProvider == null)
                    {
                        throw new Exception("No valid cert was found.");
                    }

                    byte[] requestDataBytes = EncodingWindows1251.GetBytes(requestData);

                    SHA1Managed sha1 = new SHA1Managed();
                    byte[] signedContent = rsaProvider.SignData(requestDataBytes, sha1);

                    return EncodingWindows1251.GetString(signedContent);
                }
            }
            catch
            {
                return null;
            }
        }

        // C=BG, L=Sofia, S=Sofia, O=State e-Government Agency, OU=V5400553, CN=e-gov.bg, E=rddimitrov@e-gov.bg
        private static Dictionary<string, string> GetCertificateParameters(X509Certificate2 boricaCertificate)
        {
            Dictionary<string, string> subjectByKeys = new Dictionary<string, string>();

            char[] commaSeparator = new char[] { ',' };
            char[] equalSeparator = new char[] { '=' };

            string[] keyValues = boricaCertificate.Subject
                .Split(commaSeparator, StringSplitOptions.RemoveEmptyEntries);

            foreach (string pair in keyValues)
            {
                string[] keyValue = pair.Split(equalSeparator, StringSplitOptions.RemoveEmptyEntries);

                if (keyValue.Length == 2)
                {
                    subjectByKeys.Add(keyValue[0].Trim(), keyValue[1].Trim());
                }

            }

            return subjectByKeys;
        }
    }
}