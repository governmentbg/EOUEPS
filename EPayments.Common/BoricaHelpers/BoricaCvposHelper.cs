using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace EPayments.Common.BoricaHelpers
{
    public static class BoricaCvposHelper
    {
        public static int GetTransactionOrder(string boricaTransactionOrder)
        {
            return int.Parse(boricaTransactionOrder.Substring(6));
        }

        public static string GetBoricaCvposDevPfxPath()
        {
            var pfx = Path.Combine(AppSettings.EPaymentsWeb_BoricaCentralVposCertificateFolder, AppSettings.EPaymentsWeb_CentralVposPrivateKeyFileName);
            return pfx;
        }

        public static X509Certificate2 GetBoricaCvposDev()
        {
            X509Certificate2Collection collection = new X509Certificate2Collection();
            string certificatePath = Path.Combine(AppSettings.EPaymentsWeb_BoricaCentralVposCertificateFolder, AppSettings.EPaymentsWeb_BorikaPublicKeyFileName);
            collection.Import(certificatePath);
            return collection[0];
        }
    }
}
