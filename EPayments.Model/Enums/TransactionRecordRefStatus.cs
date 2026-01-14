using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum TransactionRecordRefStatus
    {
        [Description("Друго")]
        NotReferenced = 1,

        [Description("Платено задължение")]
        ReferencedSuccessfully = 2,

        [Description("Недостатъчна сума")]
        ReferencedInsufficientAmount = 3,

        [Description("Надплатена сума")]
        ReferencedOverpaidAmount = 4,
    }
}
