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
    public class AisRequestDO
    {
        public bool IsRequestValid { get; set; }
        public EserviceClient EserviceClient { get; set; }
        public PaymentRequest PaymentRequest { get; set; }
    }
}