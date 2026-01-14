using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Controllers.v1.DataObjects
{
    internal class PaymentRequestJsonDO
    {
        public string AisPaymentId { get; set; }
        public string ServiceProviderName { get; set; }
        public string ServiceProviderBank { get; set; }
        public string ServiceProviderBIC { get; set; }
        public string ServiceProviderIBAN { get; set; }
        public string Currency { get; set; }
        public string PaymentTypeCode { get; set; }
        public string ObligationType { get; set; }
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
    }
}
