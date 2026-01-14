using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.CertificateUtils
{
    public class CertificateDO
    {
        public X509Certificate2 X509Certificate { get; set; }

        public bool IsPersonal { get; set; }

        public bool IsCompany { get; set; }

        public string PersonalIdentifier { get; set; }

        public string Name { get; set; }

        public bool HasPersonalIdentifier
        {
            get
            {
                return !String.IsNullOrWhiteSpace(this.PersonalIdentifier);
            }
        }
    }
}
