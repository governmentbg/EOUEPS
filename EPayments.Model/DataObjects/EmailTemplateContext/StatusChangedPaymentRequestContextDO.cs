using EPayments.Common.Helpers;
using EPayments.Common;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace EPayments.Model.DataObjects.EmailTemplateContext
{
    public class StatusChangedPaymentRequestContextDO
    {
        public string PaymentRequestIdentifier { get; set; }
        public string ServiceProviderName { get; set; }
        public string PaymentReason { get; set; }
        public string PaymentAmount { get; set; }
        public string Status { get; set; }
        public string PortalUrl { get; set; }
        public string FeedbackUrl { get; set; }

        public StatusChangedPaymentRequestContextDO(string paymentRequestIdentifier, string serviceProviderName, string paymentReason, decimal paymentAmount, string status, DateTime createDate)
        {
            this.PaymentRequestIdentifier = paymentRequestIdentifier;
            this.ServiceProviderName = serviceProviderName;
            this.PaymentReason = paymentReason;
            this.PaymentAmount = CurrencyHelper.GetValueForDisplayUsingCurrencyMode(paymentAmount, createDate);
            this.Status = status;
            this.PortalUrl = AppSettings.EPaymentsCommon_WebAddress;
            this.FeedbackUrl = Formatter.UriCombine(this.PortalUrl, "Home", "Feedback").ToString();
        }
    }
}