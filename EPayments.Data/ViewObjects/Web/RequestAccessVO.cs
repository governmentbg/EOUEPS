using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Web
{
    public class RequestAccessVO
    {
        public Guid PaymentRequestGid { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public DateTime PaymentReqestCreateDate { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }
        public DateTime AccessDate { get; set; }
        public string ServiceProviderName { get; set; }
        public string PaymentReason { get; set; }
        public int AccessCount { get; set; }
    }
}
