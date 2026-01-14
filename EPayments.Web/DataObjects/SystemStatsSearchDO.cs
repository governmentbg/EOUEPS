using EPayments.Model.Enums;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.DataObjects
{
    public class SystemStatsSearchDO
    {
        public int? PeriodYear { get; set; }

        public int? EserviceClientId { get; set; }

        public int? DepartmentId { get; set; }
    }
}