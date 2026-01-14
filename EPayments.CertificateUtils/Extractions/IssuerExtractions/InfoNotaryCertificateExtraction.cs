using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPayments.CertificateUtils.Extractions.IssuerExtractions
{
    internal class InfoNotaryCertificateExtraction : CertificateExtraction
    {
        private const string IssuerName = "InfoNotary";

        public override bool MatchCertificateType(X509Certificate2 x509certificate)
        {
            return x509certificate.Issuer.Contains(IssuerName);
        }

        protected override bool IsPersonal(X509Certificate2 x509certificate)
        {
            return this.MatchRegexInCertificatePolicies(x509certificate, this.IsPersonalRegex);
        }

        protected override bool IsCompany(X509Certificate2 x509certificate)
        {
            return this.MatchRegexInCertificatePolicies(x509certificate, this.IsCompanyRegex);
        }

        protected override string GetPersonalIdentifier(X509Certificate2 x509certificate)
        {
            string returnValue = String.Empty;

            X509Extension extension = x509certificate.Extensions["Subject Alternative Name"];
            if (extension != null)
            {
                if (this.EgnRegex.IsMatch(extension.Format(true)))
                {
                    Match match = this.EgnRegex.Match(extension.Format(true));
                    return match.Groups["egn"].Value.Trim();
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
                @"1.3.6.1.4.1.22144.3.1.1"
            };
        }

        protected override List<string> GetPolicyOidsForCompanyKUKEP()
        {
            return new List<string>
            {
                @"1.3.6.1.4.1.22144.3.1.2",
                @"1.3.6.1.4.1.22144.3.2.1"
            };
        }

        #region Regex expressions

        private readonly Regex EgnRegex = new Regex(@"OID.2.5.4.3.100.1.1=(?<egn>\d+)", RegexOptions.Compiled);

        private readonly Regex NameRegex = new Regex(@"CN=(?<name>.*?)\+", RegexOptions.Compiled);

        private readonly Regex IsPersonalRegex = new Regex(@"1.3.6.1.4.1.22144.1.1.1.1", RegexOptions.Compiled);

        private readonly Regex IsCompanyRegex = new Regex(@"1.3.6.1.4.1.22144.1.1.2.1", RegexOptions.Compiled);

        #endregion
    }
}
