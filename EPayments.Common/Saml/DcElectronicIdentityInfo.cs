using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Common.Saml
{
    public class DcElectronicIdentityInfo
    {
        public bool IsValid { get; set; }

        public string EGN { get; set; }

        public string GivenName { get; set; }

        public string GivenNameLat { get; set; }

        public string MiddleName { get; set; }

        public string MiddleNameLat { get; set; }

        public string FamilyName { get; set; }

        public string FamilyNameLat { get; set; }
    }
}
