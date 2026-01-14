using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPayments.CertificateUtils.Extractions.IssuerExtractions
{
    internal class BTrustCertificateExtraction : CertificateExtraction
    {
        private const string IssuerName = "B-Trust";

        public override bool MatchCertificateType(X509Certificate2 x509certificate)
        {
            return x509certificate.Issuer.Contains(IssuerName);
        }

        protected override bool IsPersonal(X509Certificate2 x509certificate)
        {
            return this.MatchRegexInCertificatePolicies(x509certificate, this.IsPersonalRegex_Policies) && IsPersonalRegex_Subject.IsMatch(x509certificate.Subject);
        }

        protected override bool IsCompany(X509Certificate2 x509certificate)
        {
            return this.MatchRegexInCertificatePolicies(x509certificate, this.IsCompanyRegex_Policies) && IsCompanyRegex_Subject.IsMatch(x509certificate.Subject);
        }

        protected override string GetPersonalIdentifier(X509Certificate2 x509certificate)
        {
            if (this.IsPersonal(x509certificate))
            {
                if (this.EgnRegex_Personal1.IsMatch(x509certificate.Subject))
                {
                    Match match = this.EgnRegex_Personal1.Match(x509certificate.Subject);
                    return match.Groups["egn"].Value.Trim();
                }
                else if (this.EgnRegex_Personal2.IsMatch(x509certificate.Subject))
                {
                    Match match = this.EgnRegex_Personal2.Match(x509certificate.Subject);
                    return match.Groups["egn"].Value.Trim();
                }
                else
                {
                    return String.Empty;
                }
            }
            else
            {
                if (this.EgnRegex_Company1.IsMatch(x509certificate.Subject))
                {
                    Match match = this.EgnRegex_Company1.Match(x509certificate.Subject);
                    return match.Groups["egn"].Value.Trim();
                }
                else if (this.EgnRegex_Company2.IsMatch(x509certificate.Subject))
                {
                    Match match = this.EgnRegex_Company2.Match(x509certificate.Subject);
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
            return new List<string>
            {
                @"1.3.6.1.4.1.15862.1.6.1.1"
            };
        }

        protected override List<string> GetPolicyOidsForCompanyKUKEP()
        {
            return new List<string>
            {
                @"1.3.6.1.4.1.15862.1.6.1.2"
            };
        }

        #region Regex expressions

        private readonly Regex EgnRegex_Personal1 = new Regex(@"OU=EGN:(?<egn>\d+)", RegexOptions.Compiled);
        private readonly Regex EgnRegex_Personal2 = new Regex(@"S=EGN:(?<egn>\d+)", RegexOptions.Compiled);
        private readonly Regex EgnRegex_Company1 = new Regex(@"S="".*EGN:(?<egn>\d+).*""", RegexOptions.Compiled);
        private readonly Regex EgnRegex_Company2 = new Regex(@"S=EGN:(?<egn>\d+)", RegexOptions.Compiled);

        private readonly Regex NameRegex = new Regex(@"CN=(?<name>.*?),", RegexOptions.Compiled);

        private readonly Regex IsPersonalRegex_Policies = new Regex(@"1.3.6.1.4.1.15862.1.5.1.1", RegexOptions.Compiled);
        private readonly Regex IsPersonalRegex_Subject = new Regex(@"OU=EGN:", RegexOptions.Compiled);

        private readonly Regex IsCompanyRegex_Policies = new Regex(@"1.3.6.1.4.1.15862.1.5.1.1", RegexOptions.Compiled);
        private readonly Regex IsCompanyRegex_Subject = new Regex(@"OU=BULSTAT:", RegexOptions.Compiled);

        #endregion
    }
}
