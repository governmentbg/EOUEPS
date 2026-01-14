using EPayments.Data.ViewObjects.Web;
using EPayments.Web.DataObjects;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Home
{
    public class SystemStatsVM
    {
        public SystemStatsSearchDO SearchDO { get; set; }

        public SystemStatsVO SystemStats { get; set; }

        public Dictionary<string, string> PeriodYearDropdownValues { get; set; }

        public Dictionary<string, string> EserviceClientIdDropdownValues { get; set; }

        public Dictionary<string, string> DepartmentIdDropdownValues { get; set; }
    }
}