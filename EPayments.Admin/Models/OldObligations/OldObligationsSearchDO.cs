using EPayments.Model.Enums;
using System;

namespace EPayments.Admin.Models.OldObligations
{
    public class OldObligationsSearchDO
    {
        public string OoDateFrom { get; set; }

        public string OoDateTo { get; set; }

        public ObligationStatusEnum? OoObligationStatus { get; set; }

        public int OoPage { get; set; } = 1;

        public OldObligationColumn OoSortBy { get; set; } = OldObligationColumn.CreateDate;

        public bool OoSortDesc { get; set; } = true;


        public bool ShowFilters
        {
            get
            {
                return
                    !String.IsNullOrWhiteSpace(this.OoDateFrom) ||
                    !String.IsNullOrWhiteSpace(this.OoDateTo) ||
                    this.OoObligationStatus.HasValue;
            }
        }

        public object ToSortRequestsRouteValues(OldObligationColumn sortBy)
        {
            return new
            {
                @ooPage = 1,
                @ooDateFrom = this.OoDateFrom,
                @ooDateTo = this.OoDateTo,
                @ooObligationStatus = this.OoObligationStatus,
                @ooSortBy = sortBy,
                @ooSortDesc = this.OoSortBy == sortBy ? !this.OoSortDesc : false,
            };
        }

        public object ToSortAllRequestsRouteValues(OldObligationColumn sortBy, bool ooSortDesc)
        {
            return new
            {
                @ooPage = 1,
                @ooDateFrom = this.OoDateFrom,
                @ooDateTo = this.OoDateTo,
                @ooObligationStatus = this.OoObligationStatus,
                @ooSortBy = sortBy,
                @ooSortDesc = ooSortDesc
            };
        }

        public object ToRequestsRouteValues()
        {
            return new
            {
                @ooDateFrom = this.OoDateFrom,
                @ooDateTo = this.OoDateTo,
                @ooObligationStatus = this.OoObligationStatus,
                @ooSortBy = this.OoSortBy,
                @ooSortDesc = this.OoSortDesc,
            };
        }
    }

    public enum OldObligationColumn
    {
        PaymentRequestIdentifier,
        CreateDate,
        PaymentAmount,
        ServiceProviderName,
        ObligationStatusId
    }
}