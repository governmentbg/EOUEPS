using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Admin.Controllers.DataObjects
{
    public class ListObligationTypeSearchDO
    {
        public int ObligationTypeId { get; set; }
        public string Name { get; set; }
        public int? IsActiveId { get; set; }
        public int Page { get; set; }
        public ObligationTypeListColumn SortBy { get; set; }
        public bool SortDesc { get; set; }

        public ListObligationTypeSearchDO()
        {
            this.Page = 1;
            this.SortBy = ObligationTypeListColumn.ObligationTypeId;
            this.SortDesc = true;
        }

        public object ToSortRouteValues(ObligationTypeListColumn sortBy)
        {
            return new
            {
                @page = 1,
                @oblName = this.Name,
                @isActiveId = this.IsActiveId,
                @sortBy = sortBy,
                @sortDesc = this.SortBy == sortBy ? !this.SortDesc : false,
            };
        }

        public object ToRouteValues()
        {
            return new
            {
                @oblName = this.Name,
                @isActiveId = this.IsActiveId,
                @sortBy = this.SortBy,
                @sortDesc = this.SortDesc
            };
        }
    }
}
