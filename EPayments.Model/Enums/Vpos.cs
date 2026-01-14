using System.ComponentModel;

namespace EPayments.Model.Enums
{
    public enum Vpos
    {
        [Description("ДСК (стар протокол)")]
        Dsk = 1,

        [Description("БОРИКА")]
        Borica = 2,

        [Description("ПИБ")]
        FiBank = 3,

        [Description("ДСК")]
        DskEcomm = 4,
    }
}
