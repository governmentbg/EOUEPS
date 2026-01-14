using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.Transactions
{
    public class TransactionPdfVM
    {
        public IList<BoricaTransactionVO> Transactions { get; set; }

        public decimal? CalculateTotalAmountInEuro { get; set; }

        public decimal? CalculateTotalFeeInEuro { get; set; }

        public decimal? CalculateTotalCommissionInEuro { get; set; }
    }
}