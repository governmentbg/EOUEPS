using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Common.Saml
{
    public class EAuthLoginDataDO
    {
        public string Egn { get; set; }
        public string Name { get; set; }
        public string ResponseStatusMessage { get; set; }
        public EAuthResponseStatus ResponseStatus { get; set; }
    }
}