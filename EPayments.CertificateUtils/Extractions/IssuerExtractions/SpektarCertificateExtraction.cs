using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPayments.CertificateUtils.Extractions.IssuerExtractions
{
    internal class SpektarCertificateExtraction : CertificateExtraction
    {
        private const string IssuerName = "Spektar";

        public override bool MatchCertificateType(X509Certificate2 x509certificate)
        {
            return x509certificate.Issuer.Contains(IssuerName);
        }

        protected override bool IsPersonal(X509Certificate2 x509certificate)
        {
            return
                this.MatchRegexInCertificatePolicies(x509certificate, this.IsPersonalRegex1) ||
                this.MatchRegexInCertificatePolicies(x509certificate, this.IsPersonalRegex2) ||
                this.MatchRegexInCertificatePolicies(x509certificate, this.IsPersonalRegex3);
        }

        protected override bool IsCompany(X509Certificate2 x509certificate)
        {
            return
                this.MatchRegexInCertificatePolicies(x509certificate, this.IsCompanyRegex1) ||
                this.MatchRegexInCertificatePolicies(x509certificate, this.IsCompanyRegex2) ||
                this.MatchRegexInCertificatePolicies(x509certificate, this.IsCompanyRegex3);
        }

        protected override string GetPersonalIdentifier(X509Certificate2 x509certificate)
        {
            if (IsPersonal(x509certificate))
            {
                if (this.EgnRegex_Personal.IsMatch(x509certificate.Subject))
                {
                    Match match = this.EgnRegex_Personal.Match(x509certificate.Subject);
                    return match.Groups["egn"].Value.Trim();
                }
                else
                {
                    return String.Empty;
                }
            }
            else
            {
                if (this.EgnRegex_Company.IsMatch(x509certificate.Subject))
                {
                    Match match = this.EgnRegex_Company.Match(x509certificate.Subject);
                    return match.Groups["egn"].Value.Trim();
                }
                else
                {
                    return String.Empty;
                }
            }
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
            return null;
        }

        protected override List<string> GetPolicyOidsForCompanyKUKEP()
        {
            return null;
        }

        #region Regex expressions

        private readonly Regex EgnRegex_Personal = new Regex(@"OU=EGNT:(?<egn>\d+)", RegexOptions.Compiled);
        private readonly Regex EgnRegex_Company = new Regex(@"T="".*EGN:(?<egn>\d+).*""", RegexOptions.Compiled);

        private readonly Regex NameRegex = new Regex(@"CN=(?<name>.*?),", RegexOptions.Compiled);

        private readonly Regex IsPersonalRegex1 = new Regex(@"1.3.6.1.4.1.18463.1.1.1.1", RegexOptions.Compiled);
        private readonly Regex IsPersonalRegex2 = new Regex(@"1.3.6.1.4.1.18463.1.1.1.2", RegexOptions.Compiled);
        private readonly Regex IsPersonalRegex3 = new Regex(@"1.3.6.1.4.1.18463.1.1.1.5", RegexOptions.Compiled);

        private readonly Regex IsCompanyRegex1 = new Regex(@"1.3.6.1.4.1.18463.1.1.1.3", RegexOptions.Compiled);
        private readonly Regex IsCompanyRegex2 = new Regex(@"1.3.6.1.4.1.18463.1.1.1.4", RegexOptions.Compiled);
        private readonly Regex IsCompanyRegex3 = new Regex(@"1.3.6.1.4.1.18463.1.1.1.6", RegexOptions.Compiled);

        #endregion
    }
}
