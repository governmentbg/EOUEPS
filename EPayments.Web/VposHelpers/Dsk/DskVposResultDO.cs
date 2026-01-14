using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.VposHelpers.Dsk
{
    public class DskVposResultDO
    {
        public string Payment_id { get; set; }
        public string Track_id { get; set; }
        public string Auth_code { get; set; }
        public string Reference_id { get; set; }
        public string Transaction_id { get; set; }
        public string Post_date { get; set; }
        public string Result { get; set; }
        public string Response_code { get; set; }
        public string Error_code { get; set; }
        public string Status { get; set; }
        public string Return_header { get; set; }
        public string Return_mac { get; set; }
    }
}