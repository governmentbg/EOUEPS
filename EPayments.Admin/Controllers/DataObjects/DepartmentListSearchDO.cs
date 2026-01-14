using EPayments.Model.Enums;

namespace EPayments.Admin.Controllers.DataObjects
{
    public class DepartmentListSearchDO
    {
        public string DepartmentName { get; set; }
        public string DepartmentUniqueIdentificationNumber { get; set; }
        public string DepartmentUnifiedBudgetClassifier { get; set; }

        public int Page { get; set; }
        public DepartmentListColumn SortBy { get; set; }
        public bool SortDesc { get; set; }

        public DepartmentListSearchDO()
        {
            this.Page = 1;
            this.SortBy = DepartmentListColumn.DepartmentId;
            this.SortDesc = true;
        }

        public object ToSortRouteValues(DepartmentListColumn sortBy)
        {
            return new
            {
                @page = 1,

                @departmentName = this.DepartmentName,
                @departmentUniqueIdentificationNumber = this.DepartmentUniqueIdentificationNumber,
                @departmentUnifiedBudgetClassifier = this.DepartmentUnifiedBudgetClassifier,
                @sortBy = sortBy,
                @sortDesc = this.SortBy == sortBy ? !this.SortDesc : false,
            };
        }

        public object ToRouteValues()
        {
            return new
            {
                @departmentName = this.DepartmentName,
                @departmentUniqueIdentificationNumber = this.DepartmentUniqueIdentificationNumber,
                @departmentUnifiedBudgetClassifier = this.DepartmentUnifiedBudgetClassifier,
                @sortBy = this.SortBy,
                @sortDesc = this.SortDesc,
            };
        }
    }
}