using EPayments.Admin.Controllers.DataObjects;
using EPayments.Admin.Models.Shared;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.Department
{
    public class DepartmentListVM
    {
        public bool IsSuperadmin { get; set; }

        public IList<DepartmentVO> DepartmentRecords { get; set; }
        public PagingVM DepartmentRecordsPagingOptions { get; set; }
        public DepartmentListSearchDO SearchDO { get; set; }
    }
}