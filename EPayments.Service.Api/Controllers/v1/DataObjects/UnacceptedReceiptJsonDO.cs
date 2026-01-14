using System;
using System.Collections.Generic;

namespace EPayments.Service.Api.Controllers.v1.DataObjects
{
    internal class UnacceptedReceiptJsonDO
    {
        public DateTime ValidationTime { get; set; }
        public List<string> Errors { get; set; }

        public UnacceptedReceiptJsonDO()
        {
            Errors = new List<string>();
        }
    }
}
