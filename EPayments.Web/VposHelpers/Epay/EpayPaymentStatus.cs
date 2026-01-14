namespace EPayments.Web.VposHelpers.Epay
{
    public enum EpayPaymentStatus
    {
        Paid = 1,
        Denied = 2,
        Expired = 3,
    }

    public static class EpayPaymentStatusExtensions
    {
        public static EserviceVposResultStatus ToEserviceVposResultStatus(this EpayPaymentStatus status)
        {
            switch (status)
            {
                case EpayPaymentStatus.Paid:
                    return EserviceVposResultStatus.Success;
                default:
                    return EserviceVposResultStatus.Failure;
            }
        }
    }
}