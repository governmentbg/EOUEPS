using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum TransactionRecordPaymentMethod
    {
        [Description("По банков път")]
        BankOrder = 1,

        [Description("Физически ПОС")]
        POS = 2,

        [Description("Виртуален ПОС")]
        VPOS = 3,
    }
}
