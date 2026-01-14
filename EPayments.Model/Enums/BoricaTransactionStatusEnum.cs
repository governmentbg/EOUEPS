using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum BoricaTransactionStatusEnum
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
