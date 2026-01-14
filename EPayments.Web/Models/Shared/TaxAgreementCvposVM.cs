using EPayments.Common.DataObjects;
using System;
using System.Linq;
using System.Text;

namespace EPayments.Web.Models.Shared
{
    public class TaxAgreementCvposVM
    {
        public Guid[] Gids { get; set; }
        public decimal PaymentAmountInBGN { get; set; }
        public decimal PaymentAmountInEuro { get; set; }

        public bool IsInternalPayment { get; set; }
        public AuthRequestDO ExternalRequestDO { get; set; }

        public string GidsToString
        {
            get
            {
                var result = Gids.Select(g => g.ToString())
               .Aggregate(new StringBuilder(),
                          (sb, str) => sb.Append("," + str),
                          sb => sb.ToString());
                return result;
            }
        }
    }
}