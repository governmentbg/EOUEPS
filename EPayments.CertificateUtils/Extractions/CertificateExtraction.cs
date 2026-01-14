using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;

namespace EPayments.CertificateUtils.Extractions
{
    internal abstract class CertificateExtraction : ICertificateExtraction
    {
        public abstract bool MatchCertificateType(X509Certificate2 x509certificate);

        protected abstract bool IsPersonal(X509Certificate2 x509certificate);

        protected abstract bool IsCompany(X509Certificate2 x509certificate);

        protected abstract string GetPersonalIdentifier(X509Certificate2 x509certificate);

        protected abstract string GetName(X509Certificate2 x509certificate);

        protected abstract List<string> GetPolicyOidsForPersonalKUKEP();

        protected abstract List<string> GetPolicyOidsForCompanyKUKEP();

        protected bool MatchRegexInCertificatePolicies(X509Certificate2 x509certificate, Regex regex)
        {
            bool returnValue = false;

            X509Extension extension = x509certificate.Extensions["Certificate Policies"];
            if (extension != null)
            {
                if (regex.IsMatch(extension.Format(true)))
                {
                    returnValue = true;
                }
            }

            return returnValue;
        }

        public CertificateDO Extract(X509Certificate2 x509certificate)
        {
            bool isPersonalKUKEP = CheckIsPersonalKUKEP(x509certificate);
            bool isCompanyKUKEP = CheckIsCompanyKUKEP(x509certificate);

            if (isPersonalKUKEP || isCompanyKUKEP)
            {
                return new CertificateDO()
                {
                    IsPersonal = isPersonalKUKEP,
                    IsCompany = isCompanyKUKEP,
                    PersonalIdentifier = GetPersonalIdentifierKUKEP(x509certificate),
                    Name = GetNameKUKEP(x509certificate),
                    X509Certificate = x509certificate
                };
            }
            else
            {
                return new CertificateDO()
                {
                    IsPersonal = IsPersonal(x509certificate),
                    IsCompany = IsCompany(x509certificate),
                    PersonalIdentifier = GetPersonalIdentifier(x509certificate),
                    Name = GetName(x509certificate),
                    X509Certificate = x509certificate
                };
            }
        }

        private bool CheckIsPersonalKUKEP(X509Certificate2 x509certificate)
        {
            List<string> checkRegexList = GetPolicyOidsForPersonalKUKEP();

            if (checkRegexList != null)
            {
                foreach (var item in checkRegexList)
                {
                    if (MatchRegexInCertificatePolicies(x509certificate, new Regex(item)))
                        return true;
                }

                return false;
            }
            else
            {
                return false;
            }
        }

        private bool CheckIsCompanyKUKEP(X509Certificate2 x509certificate)
        {
            List<string> checkRegexList = GetPolicyOidsForCompanyKUKEP();

            if (checkRegexList != null)
            {
                foreach (var item in checkRegexList)
                {
                    if (MatchRegexInCertificatePolicies(x509certificate, new Regex(item)))
                        return true;
                }

                return false;
            }
            else
            {
                return false;
            }
        }

        private string GetPersonalIdentifierKUKEP(X509Certificate2 x509certificate)
        {
            Match match = this.EgnRegex_KUKEP.Match(x509certificate.Subject);
            return match.Groups["egn"].Value.Trim();
        }

        private string GetNameKUKEP(X509Certificate2 x509certificate)
        {
            Match matchFirstName = this.FirstNameRegex_KUKEP.Match(x509certificate.Subject);
            string firstName = matchFirstName.Groups["name"].Value.Trim();

            Match matchLastName = this.LastNameRegex_KUKEP.Match(x509certificate.Subject);
            string lastName = matchLastName.Groups["name"].Value.Trim();

            if (!String.IsNullOrWhiteSpace(firstName) && !String.IsNullOrWhiteSpace(lastName))
            {
                return String.Format("{0} {1}", firstName.Trim(), lastName.Trim());
            }
            else
            {
                return null;
            }
        }

        #region KUKEP Regex

        private readonly Regex EgnRegex_KUKEP = new Regex(@"SERIALNUMBER=PNOBG-(?<egn>\d+)", RegexOptions.Compiled);
        private readonly Regex FirstNameRegex_KUKEP = new Regex(@"G=(?<name>.*?),", RegexOptions.Compiled);
        private readonly Regex LastNameRegex_KUKEP = new Regex(@"SN=(?<name>.*?),", RegexOptions.Compiled);

        #endregion
    }
}