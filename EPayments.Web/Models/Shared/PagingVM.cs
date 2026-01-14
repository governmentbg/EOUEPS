using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Data.ViewObjects;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Shared
{
    public class PagingVM
    {
        public object RouteValues { get; set; }

        public string ActionName { get; set; }

        public string ControllerName { get; set; }

        public int CurrentPageIndex { get; set; }

        public int PageSize = AppSettings.EPaymentsWeb_MaxSearchResultsPerPage;

        public int TotalItemCount { get; set; }

        public int TotalPageCount { get; private set; }

        public string PageIndexParameterName { get; set; }

        public void Calculate()
        {
            TotalPageCount = (int)Math.Ceiling(TotalItemCount / (double)PageSize);
        }
    }
}