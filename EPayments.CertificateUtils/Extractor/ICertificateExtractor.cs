using System.Security.Cryptography.X509Certificates;

namespace EPayments.CertificateUtils.Extractor
{
    public interface ICertificateExtractor
    {
        CertificateDO ExtractCertificateDO(X509Certificate2 certificate);
    }
}