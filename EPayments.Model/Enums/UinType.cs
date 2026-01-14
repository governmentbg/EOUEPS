using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum UinType
    {
        [Description("EГН")]
        Egn = 1,
        [Description("Личен номер на чужденец")]
        Lnch = 2,
        [Description("Булстат")]
        Bulstat = 3,
    }
}
