using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Web
{
    public class RequestAccessDetailsVO
    {
        public DateTime AccessDate { get; set; }
        public string IpAddress { get; set; }
        public string EbankingClientName { get; set; }
    }
}
