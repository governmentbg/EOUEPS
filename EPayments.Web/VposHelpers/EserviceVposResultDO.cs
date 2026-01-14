using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.VposHelpers
{
    public class EserviceVposResultDO
    {
        public string RequestId { get; set; }
        public string VposResultGid { get; set; }
        public string Status { get; set; }
        public string ErrorMessage { get; set; }
        public string ResultTime { get; set; }
    }
}