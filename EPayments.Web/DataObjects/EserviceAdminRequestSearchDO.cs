using EPayments.Model.Enums;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.DataObjects
{
    public class EserviceAdminRequestSearchDO
    {
        //ProcessedRequest's fields

        public string PrId { get; set; }
        public string PrRefNumber { get; set; }
        public string PrObligationType { get; set; }
        public string PrRefId { get; set; }
        public string PrDateFrom { get; set; }
        public string PrDateTo { get; set; }
        public string PrAmountFrom { get; set; }
        public string PrAmountTo { get; set; }
        public string PrProvider { get; set; }
        public string PrReason { get; set; }
        public string PrApplicantName { get; set; }
        public string PrApplicantUin { get; set; }
        public string PrPaymentType { get; set; }
        public PaymentRequestStatus? PrStatus { get; set; }

        public ObligationStatusEnum? PrObligationStatus { get; set; }

        public int PrPage { get; set; }
        public EserviceAdminRequestColumn PrSortBy { get; set; }
        public bool PrSortDesc { get; set; }

        public string Focus { get; set; }

        public bool ShowFilters
        {
            get
            {
                return
                    !String.IsNullOrWhiteSpace(this.PrId) ||
                    !String.IsNullOrWhiteSpace(this.PrRefNumber) ||
                    !String.IsNullOrWhiteSpace(this.PrObligationType) ||
                    !String.IsNullOrWhiteSpace(this.PrRefId) ||
                    !String.IsNullOrWhiteSpace(this.PrPaymentType) ||
                    !String.IsNullOrWhiteSpace(this.PrDateFrom) ||
                    !String.IsNullOrWhiteSpace(this.PrDateTo) ||
                    !String.IsNullOrWhiteSpace(this.PrAmountFrom) ||
                    !String.IsNullOrWhiteSpace(this.PrAmountTo) ||
                    !String.IsNullOrWhiteSpace(this.PrProvider) ||
                    !String.IsNullOrWhiteSpace(this.PrReason) ||
                    !String.IsNullOrWhiteSpace(this.PrApplicantName) ||
                    !String.IsNullOrWhiteSpace(this.PrApplicantUin) ||
                    this.PrStatus.HasValue ||
                    this.PrObligationStatus.HasValue;
            }
        }

        public EserviceAdminRequestSearchDO()
        {
            this.PrPage = 1;
            this.PrSortBy = EserviceAdminRequestColumn.Date;
            this.PrSortDesc = true;
        }

        public object ToSortRequestsRouteValues(EserviceAdminRequestColumn sortBy)
        {
            return new
            {
                @prPage = 1,
                @prId = this.PrId,
                @prRefNumber = this.PrRefNumber,
                @prObligationType = this.PrObligationType,
                @prRefId = this.PrRefId,
                @prPaymentType = this.PrPaymentType,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prObligationStatus = this.PrObligationStatus,

                @prApplicantName = this.PrApplicantName,
                @prApplicantUin = this.PrApplicantUin,

                @prSortBy = sortBy,
                @prSortDesc = this.PrSortBy == sortBy ? !this.PrSortDesc : false,

                //@focus = Constants.ProcessedPaymentsFocusId,
            };
        }

        public object ToSortAllRequestsRouteValues(EserviceAdminRequestColumn sortBy, bool sortDesc)
        {
            return new
            {
                @prPage = 1,
                @prId = this.PrId,
                @prRefNumber = this.PrRefNumber,
                @prObligationType = this.PrObligationType,
                @prRefId = this.PrRefId,
                @prPaymentType = this.PrPaymentType,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prObligationStatus = this.PrObligationStatus,

                @prApplicantName = this.PrApplicantName,
                @prApplicantUin = this.PrApplicantUin,

                @prSortBy = sortBy,
                @prSortDesc = sortDesc,

                //@focus = Constants.ProcessedPaymentsFocusId,
            };
        }

        public object ToRequestsRouteValues()
        {
            return new
            {
                @prId = this.PrId,
                @prRefNumber = this.PrRefNumber,
                @prObligationType = this.PrObligationType,
                @prRefId = this.PrRefId,
                @prPaymentType = this.PrPaymentType,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prObligationStatus = this.PrObligationStatus,

                @prApplicantName = this.PrApplicantName,
                @prApplicantUin = this.PrApplicantUin,

                @prSortBy = this.PrSortBy,
                @prSortDesc = this.PrSortDesc,

                //@focus = Constants.ProcessedPaymentsFocusId,
            };
        }
    }
}