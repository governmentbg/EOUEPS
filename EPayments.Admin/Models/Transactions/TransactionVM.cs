using EPayments.Admin.Models.Shared;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.Transactions
{
    public class TransactionVM
    {
        public IList<BoricaTransactionVO> Transactions { get; set; }

        public PagingVM RequestsPagingOptions { get; set; }

        public TransactionSearchDO SearchDO { get; set; }

        public decimal? CalculateTotalAmountInEuro { get; set; }

        public decimal? CalculateTotalFeeInEuro { get; set; }

        public decimal? CalculateTotalCommissionInEuro { get; set; }
    }
}