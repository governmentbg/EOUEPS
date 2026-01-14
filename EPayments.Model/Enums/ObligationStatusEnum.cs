using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum ObligationStatusEnum
    {
        [Description("Заявено")]
        Asked = 1,

        [Description("Наредено")]
        Ordered = 2,

        [Description("Неотменимо нареждане")]
        IrrevocableOrder = 3,

        [Description("Отказано плащане")]
        Canceled = 4,

        [Description("Платено по сметка на МЕУ")]
        Paid = 5,

        [Description("За разпределение")]
        ForDistribution = 6,

        [Description("Заверена сметка на администрация")]
        CheckedAccount = 7
    }
}
