using EPayments.Admin.Models.Shared;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.PaymentRequests
{
    public class PaymentRequestVM
    {
        public IList<PaymentRequestVO> Requests { get; set; }

        public PagingVM RequestsPagingOptions { get; set; }

        public PaymentRequestSearchDO SearchDO { get; set; }
    }
}