using EPayments.Common.DataObjects;
using System;

namespace EPayments.Web.Models.Shared
{
    public class TaxAgreementBoricaVM
    {
        public Guid? Gid { get; set; }
        public decimal PaymentAmountInBGN { get; set; }
        public decimal PaymentAmountInEuro { get; set; }

        public bool IsInternalPayment { get; set; }
        public AuthRequestDO ExternalRequestDO { get; set; }
    }
}