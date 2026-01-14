using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Linq;
using System;
using EPayments.CertificateUtils.Extractions;


namespace EPayments.CertificateUtils.Extractor
{
    internal class CertificateExtractor : ICertificateExtractor
    {
        private IEnumerable<ICertificateExtraction> extractions;

        public CertificateExtractor(IEnumerable<ICertificateExtraction> extractions)
        {
            this.extractions = extractions;
        }

        public CertificateDO ExtractCertificateDO(X509Certificate2 certificate)
        {
            var extraction = (CertificateExtraction)this.extractions
                .Where(e => e.MatchCertificateType(certificate))
                .SingleOrDefault();

            if (extraction != null)
            {
                return extraction.Extract(certificate);
            }
            else
            {
                throw new Exception("Certificate issuer is not recognized.");
            }
        }
    }
}