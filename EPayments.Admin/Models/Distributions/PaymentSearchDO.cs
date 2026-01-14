namespace EPayments.Admin.Models.Distributions
{
    public class PaymentSearchDO
    {
        public int Id { get; set; }

        public int CurrentPage { get; set; } = 1;

        public DistributedPaymentSortEnum SortBy { get; set; } = DistributedPaymentSortEnum.PaymentRequestIdentifier;

        public bool SortDesc { get; set; } = true;

        public string PageIndexParameterName => "CurrentPage";

        public object ToDistributionRouteValues(DistributedPaymentSortEnum sortBy)
        {
            return new
            {
                @Id = Id,
                @CurrentPage = CurrentPage,
                @SortBy = sortBy,
                @SortDesc = this.SortBy == sortBy ? !this.SortDesc : false
            };
        }

        public object ToDistributionAllRouteValues(DistributedPaymentSortEnum sortBy, bool sortDesc)
        {
            return new
            {
                @Id = Id,
                @CurrentPage = CurrentPage,
                @SortBy = sortBy,
                @SortDesc = sortDesc
            };
        }
    }

    public enum DistributedPaymentSortEnum
    {
        PaymentRequestIdentifier,
        PaymentReason,
        PaymentAmount,
        EServiceClientName,
        TargetEServiceClientName,
        ApplicantName,
        PaymentRequestStatus,
        ObligationStatus,
    }
}