using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Models;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Settings
{
    public class RequestAccessDetailsVM
    {
        public List<RequestAccessDetailsVO> AccessDetails { get; set; }

        public string PaymentRequestIdentifier { get; set; }

        public int LimitCount { get; set; }

        public int AccessCount { get; set; }

        public RequestAccessDetailsVM()
        {
            AccessDetails = new List<RequestAccessDetailsVO>();
        }
    }
}