using System;

namespace EPayments.Admin.Models.Distributions
{
    public class DistributionRevenueSearchDO
    {
        public int? DistributionRevenueId { get; set; }

        public string StartDate { get; set; }

        public string EndDate { get; set; }

        public int? DistributionType { get; set; }

        public int CurrentPage { get; set; } = 1;

        public DistributionRevenueSortEnum SortBy { get; set; } = DistributionRevenueSortEnum.CreatedAt;

        public bool SortDesc { get; set; } = true;

        public string PageIndexParameterName => "CurrentPage";

        public object ToDistributionSourtRouteValues(DistributionRevenueSortEnum sortBy)
        {
            return new
            {
                @StartDate = this.StartDate,
                @EndDate = this.EndDate,
                @DistribtuionType = this.DistributionType,
                @CurrentPage = this.CurrentPage,
                @SortBy = sortBy,
                @SortDesc = this.SortBy == sortBy ? !this.SortDesc : false
            };
        }
        public object ToDistributionRouteValues(DistributionRevenueSortEnum sortBy)
        {
            return new
            {
                @StartDate = this.StartDate,
                @EndDate = this.EndDate,
                @DistribtuionType = this.DistributionType,
                @CurrentPage = this.CurrentPage,
                @SortBy = sortBy,
                @SortDesc = this.SortDesc
            };
        }
    }

    public enum DistributionRevenueSortEnum
    {
        CreatedAt,
        DistributedDate,
        IsDistributed,
        Refid,
        TotalSum,
        IsFileGenerated,
        DistributionType,
        FileName
    }
}