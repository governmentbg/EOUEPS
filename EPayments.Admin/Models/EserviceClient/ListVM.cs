using EPayments.Admin.DataObjects;
using EPayments.Admin.Models.Shared;
using EPayments.Data.ViewObjects.Web;
using System.Collections.Generic;

namespace EPayments.Admin.Models.EserviceClient
{
    public class ListVM
    {
        public bool IsSuperadmin { get; set; }
        public IList<EserviceRecordVO> EserviceRecords { get; set; }
        public PagingVM EserviceRecordsPagingOptions { get; set; }
        public ListSearchDO SearchDO { get; set; }
        public Dictionary<int, string> BoricaWarnings { get; set; }
    }
}