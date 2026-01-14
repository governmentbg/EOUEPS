using System;
namespace EPayments.Model.DataObjects
{
    public class NotificationPostDataJsonAlgorithm2: NotificationPostDataJson
    {
        public string AisPaymentId { get; set; }

        public string PaymentReferenceNumber { get; set; }

        public string PaymentAmount { get; set; }
    }
}
