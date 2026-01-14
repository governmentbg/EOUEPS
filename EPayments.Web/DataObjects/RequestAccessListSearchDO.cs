using EPayments.Model.Enums;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.DataObjects
{
    public class RequestAccessListSearchDO
    {
        //PendingRequest's fields

        public int Page { get; set; }
        public RequestAccessListColumn SortBy { get; set; }
        public bool SortDesc { get; set; }

        public RequestAccessListSearchDO()
        {
            this.Page = 1;
            this.SortBy = RequestAccessListColumn.AccessDate;
            this.SortDesc = true;
        }

        public object ToSortRequestAccessListRouteValues(RequestAccessListColumn sortBy)
        {
            return new
            {
                @page = 1,
                @sortBy = sortBy,
                @sortDesc = this.SortBy == sortBy ? !this.SortDesc : false,
            };
        }

        public object ToRequestAccessListRouteValues()
        {
            return new
            {
                @sortBy = this.SortBy,
                @sortDesc = this.SortDesc,
            };
        }
    }
}