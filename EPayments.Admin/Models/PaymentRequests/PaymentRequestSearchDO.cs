using EPayments.Model.Enums;
using System;

namespace EPayments.Admin.Models.PaymentRequests
{
    public class PaymentRequestSearchDO
    {
        public string PrId { get; set; }

        public string PrRefenceNumber { get; set; }

        public string PrDateFrom { get; set; }

        public string PrDateTo { get; set; }

        public string PrAmountFrom { get; set; }

        public string PrAmountTo { get; set; }

        public string PrProvider { get; set; }

        public string PrReason { get; set; }

        public string PrApplicantName { get; set; }

        public string PrApplicantUin { get; set; }

        public PaymentRequestStatus? PrPaymentStatus { get; set; }

        public PaymentRequestStatus? PrPaymentStatusChanged { get; set; }

        public ObligationStatusEnum? PrObligationStatus { get; set; }

        public int PrPage { get; set; } = 1;

        public PaymentRequestColumn PrSortBy { get; set; } = PaymentRequestColumn.CreateDate;

        public bool PrSortDesc { get; set; } = true;

        public string Focus { get; set; }

        public bool IsSearchForm { get; set; } = true;

        public bool ShowFilters
        {
            get
            {
                return
                    !String.IsNullOrWhiteSpace(this.PrId) ||
                    !String.IsNullOrWhiteSpace(this.PrDateFrom) ||
                    !String.IsNullOrWhiteSpace(this.PrDateTo) ||
                    !String.IsNullOrWhiteSpace(this.PrAmountFrom) ||
                    !String.IsNullOrWhiteSpace(this.PrAmountTo) ||
                    !String.IsNullOrWhiteSpace(this.PrProvider) ||
                    !String.IsNullOrWhiteSpace(this.PrReason) ||
                    !String.IsNullOrWhiteSpace(this.PrApplicantName) ||
                    !String.IsNullOrWhiteSpace(this.PrApplicantUin) ||
                    this.PrPaymentStatus.HasValue;
            }
        }

        public object ToSortRequestsRouteValues(PaymentRequestColumn sortBy)
        {
            return new
            {
                @prPage = 1,
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prPaymentStatus = this.PrPaymentStatus,
                @prObligationStatus = this.PrObligationStatus,
                @prApplicantName = this.PrApplicantName,
                @prApplicantUin = this.PrApplicantUin,
                @prSortBy = sortBy,
                @prSortDesc = this.PrSortBy == sortBy ? !this.PrSortDesc : false,
            };
        }

        public object ToSortAllRequestsRouteValues(PaymentRequestColumn sortBy, bool sortDesc)
        {
            return new
            {
                @prPage = 1,
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prPaymentStatus = this.PrPaymentStatus,
                @prObligationStatus = this.PrObligationStatus,
                @prApplicantName = this.PrApplicantName,
                @prApplicantUin = this.PrApplicantUin,
                @prSortBy = sortBy,
                @prSortDesc = sortDesc,
            };
        }

        public object ToRequestsRouteValues()
        {
            return new
            {
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prPaymentStatus = this.PrPaymentStatus,
                @prObligationStatus = this.PrObligationStatus,
                @prApplicantName = this.PrApplicantName,
                @prApplicantUin = this.PrApplicantUin,
                @prSortBy = this.PrSortBy,
                @prSortDesc = this.PrSortDesc,
            };
        }
    }

    public enum PaymentRequestColumn
    {
        CreateDate,
        ApplicantName,
        PaymentRequestIdentifier,
        PaymentReferenceNumber,
        ExpirationDate,
        ServiceProviderName,
        PaymentReason,
        PaymentAmount,
        PaymentRequestStatusId,
        ObligationStatusId
    }
}