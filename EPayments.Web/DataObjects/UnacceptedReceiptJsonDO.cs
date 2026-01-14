using System;
using System.Collections.Generic;

namespace EPayments.Web.DataObjects
{
    public class UnacceptedReceiptJsonDO
    {
        public DateTime ValidationTime { get; set; }

        public List<string> Errors { get; set; }
    }
}