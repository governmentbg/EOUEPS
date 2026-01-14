using EPayments.Data.ViewObjects.Web;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Shared;
using System.Collections.Generic;

namespace EPayments.Web.Models.Payments
{
    public class ListVM
    {
        public IList<PendingRequestVO> PendingPayments { get; set; }
        public IList<ProcessedRequestVO> ProcessedPayments { get; set; }
        public PagingVM PendingPaymentsPagingOptions { get; set; }
        public PagingVM ProcessedPaymentsPagingOptions { get; set; }
        public ListSearchDO SearchDO { get; set; }
        public decimal TotalAmmount { get; set; }
        public List<string> DisabledItems { get; set; }

    }
}