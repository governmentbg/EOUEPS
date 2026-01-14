using EPayments.Admin.Models.Shared;
using EPayments.Common.Helpers;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.UndistributedPayments
{
    public class UndistributedPaymentVM
    {
        public IList<UndistributedPaymentRequestVO> Requests { get; set; }

        public PagingVM RequestsPagingOptions { get; set; }

        public UndistributedPaymentSearchDO SearchDO { get; set; }

        public decimal TotalInEuro 
        { 
            get
            {
                if (Requests.Count == 0)
                    return 0;
                else
                {
                    decimal total = 0;
                    foreach (var request in Requests)
                    {
                        total += CurrencyHelper.GetEuroValue(request.PaymentAmount, request.CreateDate);
                    }
                    return total;
                }
            } 
        }
    }
}