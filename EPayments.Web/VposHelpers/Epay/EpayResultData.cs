namespace EPayments.Web.VposHelpers.Epay
{
    public class EpayResultData
    {
        public string Invoice { get; set; }
        public EpayPaymentStatus Status { get; set; }
    }
}