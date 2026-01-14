namespace EPayments.Data.ViewObjects.Admin
{
    public class TotalsBoricaTransactionVO
    {
        public int TotalPages { get; set; }

        public decimal TotalAmountInEuro { get; set; }

        public decimal TotalFeeInEuro { get; set; }

        public decimal CommissionInEuro { get; set; }
    }
}
