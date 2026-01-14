using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Api
{
    public class RequestXmlVO
    {
        public string Id { get; set; }
        public string RequestXml { get; set; }
        public PaymentRequest PaymentRequest { get; set; }
    }
}
