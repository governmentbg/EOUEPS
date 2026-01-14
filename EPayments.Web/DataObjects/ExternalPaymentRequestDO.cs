using EPayments.Model.Enums;
using EPayments.Model.Models;
using EPayments.Web.Common;
using EPayments.Web.VposHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.DataObjects
{
    public class ExternalPaymentRequestDO
    {
        public bool IsRequestValid { get; set; }
        public EserviceVposResultStatus? ErrorStatus { get; set; }
        public string OkUrl { get; set; }
        public string CancelUrl { get; set; }
        public EserviceClient EserviceClient { get; set; }
        public PaymentRequest PaymentRequest { get; set; }
    }
}