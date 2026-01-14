using EPayments.Model.Enums;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.DataObjects
{
    public class ListSearchDO
    {
        //PendingRequest's fields

        public int PPage { get; set; }
        public PendingPaymentColumn PSortBy { get; set; }
        public bool PSortDesc { get; set; }

        //ProcessedRequest's fields

        public string PrId { get; set; }
        public string PrDateFrom { get; set; }
        public string PrDateTo { get; set; }
        public string PrAmountFrom { get; set; }
        public string PrAmountTo { get; set; }
        public string PrProvider { get; set; }
        public string PrReason { get; set; }
        public PaymentRequestStatus? PrStatus { get; set; }

        public int PrPage { get; set; }
        public ProcessedPaymentColumn PrSortBy { get; set; }
        public bool PrSortDesc { get; set; }

        public string Focus { get; set; }

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
                    this.PrStatus.HasValue;
            }
        }

        public ListSearchDO()
        {
            this.PPage = 1;
            this.PSortBy = PendingPaymentColumn.PaymentId;
            this.PSortDesc = true;

            this.PrPage = 1;
            this.PrSortBy = ProcessedPaymentColumn.Date;
            this.PrSortDesc = true;
        }

        public object ToSortPendingRequestsRouteValues(PendingPaymentColumn sortBy)
        {
            return new
            {
                @pPage = 1,
                @pSortBy = sortBy,
                @pSortDesc = this.PSortBy == sortBy ? !this.PSortDesc : false,

                @prPage = this.PrPage,
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prSortBy = this.PrSortBy,
                @prSortDesc = this.PrSortDesc,

                //@focus = (string)null,
            };
        }

        public object ToSortProcessedRequestsRouteValues(ProcessedPaymentColumn sortBy)
        {
            return new
            {
                @pPage = this.PPage,
                @pSortBy = this.PSortBy,
                @pSortDesc = this.PSortDesc,

                @prPage = 1,
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prSortBy = sortBy,
                @prSortDesc = this.PrSortBy == sortBy ? !this.PrSortDesc : false,

                @focus = Constants.ProcessedPaymentsFocusId,
            };
        }

        public object ToSortAllProcessedRequestsRouteValues(ProcessedPaymentColumn sortBy, bool sortDesc)
        {
            return new
            {
                @pPage = this.PPage,
                @pSortBy = this.PSortBy,
                @pSortDesc = this.PSortDesc,

                @prPage = 1,
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prSortBy = sortBy,
                @prSortDesc = sortDesc,

                @focus = Constants.ProcessedPaymentsFocusId,
            };
        }

        public object ToPendingRequestsRouteValues()
        {
            return new
            {
                @pSortBy = this.PSortBy,
                @pSortDesc = this.PSortDesc,

                @prPage = this.PrPage,
                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prSortBy = this.PrSortBy,
                @prSortDesc = this.PrSortDesc,

                //@focus = (string)null,
            };
        }

        public object ToProcessedRequestsRouteValues()
        {
            return new
            {
                @pPage = this.PPage,
                @pSortBy = this.PSortBy,
                @pSortDesc = this.PSortDesc,

                @prId = this.PrId,
                @prDateFrom = this.PrDateFrom,
                @prDateTo = this.PrDateTo,
                @prAmountFrom = this.PrAmountFrom,
                @prAmountTo = this.PrAmountTo,
                @prProvider = this.PrProvider,
                @prReason = this.PrReason,
                @prStatus = this.PrStatus,
                @prSortBy = this.PrSortBy,
                @prSortDesc = this.PrSortDesc,

                @focus = Constants.ProcessedPaymentsFocusId,
            };
        }
    }
}