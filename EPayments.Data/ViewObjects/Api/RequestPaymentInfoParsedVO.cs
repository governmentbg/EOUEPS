using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Api
{
    public class RequestPaymentInfoParsedVO
    {
        public string Id { get; set; }
        public PaymentRequestStatus? Status { get; set; }
        public DateTime? StatusChangeTime { get; set; }
        public string ServiceProviderName { get; set; }
        public string ServiceProviderBank { get; set; }
        public string ServiceProviderBIC { get; set; }
        public string ServiceProviderIBAN { get; set; }
        public string Currency { get; set; }
        public string PaymentTypeCode { get; set; }
        public decimal? PaymentAmount { get; set; }
        public string PaymentReason { get; set; }
        public UinType? ApplicantUinType { get; set; }
        public string ApplicantUin { get; set; }
        public string ApplicantName { get; set; }
        public string PaymentReferenceType { get; set; }
        public string PaymentReferenceNumber { get; set; }
        public DateTime? PaymentReferenceDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string AdditionalInformation { get; set; }
        public DateTime? CreateDate { get; set; }
        public string EserviceClientAisName { get; set; }
        public string ClientId { get; set; }
        public string EserviceClientServiceName { get; set; }
        public string EserviceClientDepartmentId { get; set; }
    }
}
