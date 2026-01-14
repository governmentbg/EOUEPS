using System.ComponentModel;

namespace EPayments.Admin.Models.Transactions
{
    public class TransactionSearchDO
    {
        public string TtDateFrom { get; set; }

        public string TtDateTo { get; set; }

        public TransactionStatus? TtTransactionStatus { get; set; }

        public int TtPage { get; set; } = 1;

        public TransactionColumn TtSortBy { get; set; } = TransactionColumn.TransactionDate;

        public bool TtSortDesc { get; set; } = true;

        public object ToSortRequestsRouteValues(TransactionColumn sortBy)
        {
            return new
            {
                @TtDateFrom = this.TtDateFrom,
                @TtDateTo = this.TtDateTo,
                @TtTransactionStatus = this.TtTransactionStatus,
                @TtSortBy = sortBy,
                @TtSortDesc = this.TtSortBy == sortBy ? !this.TtSortDesc : false
            };
        }

        public object ToSortAllRequestsRouteValues(TransactionColumn sortBy, bool ooSortDesc)
        {
            return new
            {
                @TtPage = 1,
                @TtDateFrom = this.TtDateFrom,
                @TtDateTo = this.TtDateTo,
                @TtTransactionStatus = this.TtTransactionStatus,
                @TtSortBy = sortBy,
                @TtSortDesc = ooSortDesc
            };
        }

        public object ToRequestsRouteValues()
        {
            return new
            {
                @TtDateFrom = this.TtDateFrom,
                @TtDateTo = this.TtDateTo,
                @TtTransactionStatus = this.TtTransactionStatus,
                @TtSortBy = this.TtSortBy,
                @TtSortDesc = this.TtSortDesc,
            };
        }
    }

    public enum TransactionColumn
    {
        [Description("Номер на поръчка")]
        Order,
        [Description("Сума")]
        Amount,
        [Description("Такса")]
        Fee,
        [Description("Комисионна")]
        Commission,
        [Description("Дата на транзакцията")]
        TransactionDate,
        [Description("Номер на карта")]
        Card,
        [Description("Дата на стълмент")]
        SettlementDate,
        [Description("Съобщение от Борика")]
        StatusMessage
    }

    public enum TransactionStatus
    {
        [Description("Очаква плащане")]
        Pending = 1,
        [Description("Платена")]
        Paid = 2,
        [Description("Получена такса")]
        TaxReceived = 3,
        [Description("Разпределена")]
        Distributed = 4,
        [Description("Отказана")]
        Canceled = 5
    }
}