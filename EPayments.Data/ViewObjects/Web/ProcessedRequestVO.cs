using EPayments.Model.Enums;
using System;
using System.Collections.Generic;

namespace EPayments.Data.ViewObjects.Web
{
    public class ProcessedRequestVO
    {
        public Guid Gid { get; set; }
        public string ApplicantName { get; set; }
        public string ApplicantUin { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public string PaymentReferenceNumber { get; set; }
        public DateTime? TransactionDate { get; set; }
        public string ServiceProviderName { get; set; }
        public string PaymentReason { get; set; }
        public decimal PaymentAmountRequest { get; set; }
        public string AdditionalInformation { get; set; }
        public Guid PaymentRequestGid { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }
        public DateTime ExpirationDate { get; set; }
        public DateTime CreateDate { get; set; }
        public string ObligationType { get; set; }
        public string PaymentReferenceType { get; set; }
        public ObligationStatusEnum? ObligationStatusId { get; set; }
        public IEnumerable<PaymentRequestObligationLog> PaymentRequestObligationLogs { get; set; }
        public string InitiatorName { get; set; }
        public string Refid { get; set; }
    }
}
