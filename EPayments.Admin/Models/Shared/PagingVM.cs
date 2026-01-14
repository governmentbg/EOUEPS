using EPayments.Common;
using System;

namespace EPayments.Admin.Models.Shared
{
    public class PagingVM
    {
        public object RouteValues { get; set; }

        public string ActionName { get; set; }

        public string ControllerName { get; set; }

        public int CurrentPageIndex { get; set; }

        public int PageSize { get; set; } = AppSettings.EPaymentsWeb_MaxSearchResultsPerPage;

        public int TotalItemCount { get; set; }

        public int TotalPageCount { get; private set; }

        public string PageIndexParameterName { get; set; }

        public void Calculate()
        {
            TotalPageCount = (int)Math.Ceiling(TotalItemCount / (double)PageSize);
        }
    }
}