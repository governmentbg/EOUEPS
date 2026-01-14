using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPayments.CertificateUtils.Extractions.IssuerExtractions
{
    internal class EvrotrustCertificateExtraction : CertificateExtraction
    {
        private const string IssuerName = "Evrotrust";

        public override bool MatchCertificateType(X509Certificate2 x509certificate)
        {
            return x509certificate.Issuer.Contains(IssuerName);
        }

        protected override bool IsPersonal(X509Certificate2 x509certificate)
        {
            return !x509certificate.Subject.Contains(CompanyIdentifier);
        }

        protected override bool IsCompany(X509Certificate2 x509certificate)
        {
            return x509certificate.Subject.Contains(CompanyIdentifier);
        }

        protected override string GetPersonalIdentifier(X509Certificate2 x509certificate)
        {
            string returnValue = String.Empty;

            if (this.EgnContainerRegex.IsMatch(x509certificate.Subject))
            {
                Match match = this.EgnContainerRegex.Match(x509certificate.Subject);
                string egnContainer = match.Groups["egnContainer"].Value.Trim();

                if (!String.IsNullOrWhiteSpace(egnContainer) && egnContainer.StartsWith("PNOBG-"))
                {
                    returnValue = egnContainer.Substring("PNOBG-".Length).Trim();
                }
            }

            return returnValue;
        }

        protected override string GetName(X509Certificate2 x509certificate)
        {
            if (this.NameRegex.IsMatch(x509certificate.Subject))
            {
                Match match = this.NameRegex.Match(x509certificate.Subject);
                return match.Groups["name"].Value.Trim();
            }
            else
            {
                return String.Empty;
            }
        }

        protected override List<string> GetPolicyOidsForPersonalKUKEP()
        {
            return new List<string>
            {
                @"1.3.6.1.4.1.47272.2.7"
            };
        }

        protected override List<string> GetPolicyOidsForCompanyKUKEP()
        {
            return new List<string>
            {
                @"1.3.6.1.4.1.47272.2.8"
            };
        }

        private const string CompanyIdentifier = "2.5.4.97=";

        #region Regex expressions

        private readonly Regex NameRegex = new Regex(@"CN=(?<name>.*?)\,", RegexOptions.Compiled);

        private readonly Regex EgnContainerRegex = new Regex(@"SERIALNUMBER=(?<egnContainer>.*?)\,", RegexOptions.Compiled);

        #endregion
    }
}
