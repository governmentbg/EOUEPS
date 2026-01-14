using EPayments.Model.Enums;
using System;

namespace EPayments.Admin.Models.UndistributedPayments
{
    public class UndistributedPaymentSearchDO
    {
        public string UpId { get; set; }

        public string UpDateFrom { get; set; }

        public string UpDateTo { get; set; }

        public string UpAmountFrom { get; set; }

        public string UpAmountTo { get; set; }

        public string UpProvider { get; set; }

        public string UpReason { get; set; }

        public ObligationStatusEnum? UpObligationStatus => ObligationStatusEnum.ForDistribution;

        public int UpPage { get; set; } = 1;

        public UndistributetPaymentColumn UpSortBy { get; set; } = UndistributetPaymentColumn.CreateDate;

        public bool UpSortDesc { get; set; } = true;


        public bool ShowFilters
        {
            get
            {
                return
                    !String.IsNullOrWhiteSpace(this.UpId) ||
                    !String.IsNullOrWhiteSpace(this.UpDateFrom) ||
                    !String.IsNullOrWhiteSpace(this.UpDateTo) ||
                    !String.IsNullOrWhiteSpace(this.UpAmountFrom) ||
                    !String.IsNullOrWhiteSpace(this.UpAmountTo) ||
                    !String.IsNullOrWhiteSpace(this.UpProvider) ||
                    !String.IsNullOrWhiteSpace(this.UpReason) ||
                    this.UpObligationStatus.HasValue;
            }
        }

        public object ToSortRequestsRouteValues(UndistributetPaymentColumn sortBy)
        {
            return new
            {
                @upPage = 1,
                @upId = this.UpId,
                @upDateFrom = this.UpDateFrom,
                @upDateTo = this.UpDateTo,
                @upAmountFrom = this.UpAmountFrom,
                @upAmountTo = this.UpAmountTo,
                @upProvider = this.UpProvider,
                @upReason = this.UpReason,
                @upObligationStatus = this.UpObligationStatus,
                @upSortBy = sortBy,
                @upSortDesc = this.UpSortBy == sortBy ? !this.UpSortDesc : false,
            };
        }

        public object ToSortAllRequestsRouteValues(UndistributetPaymentColumn sortBy, bool upSortDesc)
        {
            return new
            {
                @upPage = 1,
                @upId = this.UpId,
                @upDateFrom = this.UpDateFrom,
                @upDateTo = this.UpDateTo,
                @upAmountFrom = this.UpAmountFrom,
                @upAmountTo = this.UpAmountTo,
                @upProvider = this.UpProvider,
                @upReason = this.UpReason,
                @upObligationStatus = this.UpObligationStatus,
                @upSortBy = sortBy,
                @upSortDesc = upSortDesc,
            };
        }

        public object ToRequestsRouteValues()
        {
            return new
            {
                @upId = this.UpId,
                @upDateFrom = this.UpDateFrom,
                @upDateTo = this.UpDateTo,
                @upAmountFrom = this.UpAmountFrom,
                @upAmountTo = this.UpAmountTo,
                @upProvider = this.UpProvider,
                @upReason = this.UpReason,
                @upObligationStatus = this.UpObligationStatus,
                @upSortBy = this.UpSortBy,
                @upSortDesc = this.UpSortDesc,
            };
        }
    }

    public enum UndistributetPaymentColumn
    {
        CreateDate,
        PaymentRequestIdentifier,
        ServiceProviderName,
        PaymentReason,
        PaymentAmount,
        ObligationStatusId
    }
}