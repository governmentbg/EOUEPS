using EPayments.Data.ViewObjects.Web;
using System.Collections.Generic;
using EPayments.Admin.Models.Shared;
using EPayments.Admin.DataObjects;
using EPayments.Admin.Controllers.DataObjects;

namespace EPayments.Admin.Models.ObligationType
{
    public class ObligationTypeLIstVM
    {
        public bool IsSuperadmin { get; set; }
        public IList<ObligationTypeVO> ObligationTypeRecords { get; set; }
        public PagingVM ObligationTypeRecordsPagingOptions { get; set; }
        public ListObligationTypeSearchDO SearchDO { get; set; }
    }
}