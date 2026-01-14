using EPayments.Common.Helpers;
using EPayments.Common;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace EPayments.Model.DataObjects.EmailTemplateContext
{
    public class NewPaymentRequestContextDO
    {
        public string PaymentRequestIdentifier { get; set; }
        public string ServiceProviderName { get; set; }
        public string PaymentReason { get; set; }
        public string PaymentAmount { get; set; }
        public string PortalUrl { get; set; }
        public string FeedbackUrl { get; set; }
        public string AccessCode { get; set; }

        public NewPaymentRequestContextDO(string paymentRequestIdentifier, string serviceProviderName, string paymentReason, decimal paymentAmount, string accessCode, DateTime createDate)
        {
            this.PaymentRequestIdentifier = paymentRequestIdentifier;
            this.ServiceProviderName = serviceProviderName;
            this.PaymentReason = paymentReason;
            this.PaymentAmount = CurrencyHelper.GetValueForDisplayUsingCurrencyMode(paymentAmount, createDate);
            this.PortalUrl = AppSettings.EPaymentsCommon_WebAddress;
            this.FeedbackUrl = Formatter.UriCombine(this.PortalUrl, "Home", "Feedback").ToString();
            this.AccessCode = accessCode;
        }
    }
}