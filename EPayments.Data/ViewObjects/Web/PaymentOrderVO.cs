using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Web
{
    public class PaymentOrderVO
    {
        public Guid Gid { get; set; }
        public string ServiceProviderName { get; set; }
        public string BankName { get; set; }
        public string IBAN { get; set; }
        public string BIC { get; set; }
        public decimal PaymentAmount { get; set; }
        public string PaymentReason { get; set; }
        public string DocumentType { get; set; }
        public string DocumentNumber { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime DocumentDate { get; set; }
        public DateTime? PaymentPeriodStartDate { get; set; }
        public DateTime? PaymentPeriodEndDate { get; set; }
        public string ApplicantName { get; set; }
        public UinType ApplicantUinTypeId { get; set; }
        public string ApplicantUin { get; set; }
        public bool IsVposAuthorized { get; set; }
        public string VposAuthorizationId { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public DateTime PaymentRequestStatusChangeTime { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }
        public string ObligationType { get; set; }
        public string ObligationTypeCode { get; set; }
        public int ObligationTypeAlgorithmId { get; set; }
        public string AdditionalInformation { get; set; }
    }
}
