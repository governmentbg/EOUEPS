using EPayments.Model.Enums;

namespace EPayments.Web.DataObjects
{
    public class TransactionListSearchDO
    {
        public int? EserviceBankAccountId { get; set; }
        public string TransactionAccountingDateFrom { get; set; }
        public string TransactionAccountingDateTo { get; set; }
        public string TransactionAmountFrom { get; set; }
        public string TransactionAmountTo { get; set; }
        public string InfoDocumentNumber { get; set; }
        public string InfoDocumentDateFrom { get; set; }
        public string InfoDocumentDateTo { get; set; }
        public string InfoSenderIban { get; set; }
        public string InfoSenderName { get; set; }
        public string InfoDebtorName { get; set; }
        public string InfoDebtorBulstatEgnLnch { get; set; }
        public string InfoPaymentReason { get; set; }
        public string InfoAC1AuthorizationCode { get; set; }
        public int? TransactionRecordPaymentMethod { get; set; }
        public int? TransactionRecordRefStatus { get; set; }

        public int Page { get; set; }
        public TransactionListColumn SortBy { get; set; }
        public bool SortDesc { get; set; }

        public TransactionListSearchDO()
        {
            this.Page = 1;
            this.SortBy = TransactionListColumn.TransactionAccountingDate;
            this.SortDesc = true;
        }

        public object ToSortRouteValues(TransactionListColumn sortBy)
        {
            return new
            {
                @page = 1,

                @eserviceBankAccountId = this.EserviceBankAccountId,
                @transactionAccountingDateFrom = this.TransactionAccountingDateFrom,
                @transactionAccountingDateTo = this.TransactionAccountingDateTo,
                @transactionAmountFrom = this.TransactionAmountFrom,
                @transactionAmountTo = this.TransactionAmountTo,
                @infoDocumentNumber = this.InfoDocumentNumber,
                @infoDocumentDateFrom = this.InfoDocumentDateFrom,
                @infoDocumentDateTo = this.InfoDocumentDateTo,
                @infoSenderIban = this.InfoSenderIban,
                @infoSenderName = this.InfoSenderName,
                @infoDebtorName = this.InfoDebtorName,
                @infoDebtorBulstatEgnLnch = this.InfoDebtorBulstatEgnLnch,
                @infoPaymentReason = this.InfoPaymentReason,
                @infoAC1AuthorizationCode = this.InfoAC1AuthorizationCode,
                @transactionRecordPaymentMethod = this.TransactionRecordPaymentMethod,
                @transactionRecordRefStatus = this.TransactionRecordRefStatus,

                @sortBy = sortBy,
                @sortDesc = this.SortBy == sortBy ? !this.SortDesc : false,
            };
        }

        public object ToRouteValues()
        {
            return new
            {
                @eserviceBankAccountId = this.EserviceBankAccountId,
                @transactionAccountingDateFrom = this.TransactionAccountingDateFrom,
                @transactionAccountingDateTo = this.TransactionAccountingDateTo,
                @transactionAmountFrom = this.TransactionAmountFrom,
                @transactionAmountTo = this.TransactionAmountTo,
                @infoDocumentNumber = this.InfoDocumentNumber,
                @infoDocumentDateFrom = this.InfoDocumentDateFrom,
                @infoDocumentDateTo = this.InfoDocumentDateTo,
                @infoSenderIban = this.InfoSenderIban,
                @infoSenderName = this.InfoSenderName,
                @infoDebtorName = this.InfoDebtorName,
                @infoDebtorBulstatEgnLnch = this.InfoDebtorBulstatEgnLnch,
                @infoPaymentReason = this.InfoPaymentReason,
                @infoAC1AuthorizationCode = this.InfoAC1AuthorizationCode,
                @transactionRecordPaymentMethod = this.TransactionRecordPaymentMethod,
                @transactionRecordRefStatus = this.TransactionRecordRefStatus,

                @sortBy = this.SortBy,
                @sortDesc = this.SortDesc,
            };
        }
    }
}