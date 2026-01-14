using EPayments.Model.Enums;

namespace EPayments.Admin.DataObjects
{
    public class ListSearchDO
    {
        public string AisName { get; set; }
        public string DepartmentName { get; set; }
        public int? VposClientId { get; set; }
        public int? IsActiveId { get; set; }

        public int Page { get; set; }
        public EserviceListColumn SortBy { get; set; }
        public bool SortDesc { get; set; }

        public ListSearchDO()
        {
            this.Page = 1;
            this.SortBy = EserviceListColumn.EserviceClientId;
            this.SortDesc = true;
        }

        public object ToSortRouteValues(EserviceListColumn sortBy)
        {
            return new
            {
                @page = 1,

                @aisName = this.AisName,
                @departmentName = this.DepartmentName,
                @vposClientId = this.VposClientId,
                @isActiveId = this.IsActiveId,

                @sortBy = sortBy,
                @sortDesc = this.SortBy == sortBy ? !this.SortDesc : false,
            };
        }

        public object ToRouteValues()
        {
            return new
            {
                @aisName = this.AisName,
                @departmentName = this.DepartmentName,
                @vposClientId = this.VposClientId,
                @isActiveId = this.IsActiveId,

                @sortBy = this.SortBy,
                @sortDesc = this.SortDesc,
            };
        }
    }
}