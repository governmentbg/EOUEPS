using System.Security.Cryptography.X509Certificates;

namespace EPayments.CertificateUtils.Extractions
{
    internal interface ICertificateExtraction
    {
        bool MatchCertificateType(X509Certificate2 x509certificate);
    }
}
