using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.Common.HelpDesk
{
    public class ActionResultDO
    {
        public ActionResultCode ResultCode { get; set; }
        public Dictionary<string, List<string>> Errors { get; set; }
        public UnitSimpleIDO Data { get; set; }
    }
}