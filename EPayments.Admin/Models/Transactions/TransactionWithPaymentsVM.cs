using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;

namespace EPayments.Admin.Models.Transactions
{
    public class TransactionWithPaymentsVM
    {
        public BoricaTransactionVO BoricaTransaction { get; set; }

        public List<PaymentRequestVO> Payments { get; set; }
    }
}