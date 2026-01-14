using System;

namespace EPayments.Service.Api.Controllers.v1.DataObjects
{
    internal class PaymentRequestJsonExtendedDO
    {
        public string AisPaymentId { get; set; }
        public string Currency { get; set; }
        public string PaymentTypeCode { get; set; }
        public double? PaymentAmount { get; set; }
        public string PaymentReason { get; set; }
        public int? ApplicantUinTypeId { get; set; }
        public string ApplicantUin { get; set; }
        public string ApplicantName { get; set; }
        public string PaymentReferenceType { get; set; }
        public string PaymentReferenceNumber { get; set; }
        public DateTime? PaymentReferenceDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string AdditionalInformation { get; set; }
        public string AdministrativeServiceUri { get; set; }
        public string AdministrativeServiceSupplierUri { get; set; }
        public string AdministrativeServiceNotificationURL { get; set; }
        public int? PayOrder { get; set; }

        public string ObligationType { get; set; }
    }
}
