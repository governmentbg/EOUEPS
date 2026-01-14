using EPayments.Model.DataObjects;

namespace EPayments.Web.DataObjects
{
    public class MDT_ExtendedPaymentRequestDO
    {
        public MDT_ExtendedInfoJson mDT_ExtendedInfoJson { get; set; }

        public int ObligationTypeId { get; set; }

        public int? PayOrder { get; set; }

        public string PaymentRequestIdentifier { get; set; }

        public int EserviceClientId { get; set; }
    }
}
