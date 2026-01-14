using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum PaymentReferenceType
    {
        [Description("1 - Декларация")]
        Declaration = 1,

        [Description("2 - Ревизионен акт")]
        RevisionAct = 2,

        [Description("3 - Наказателно постановление")]
        PenalDecree = 3,

        [Description("4 - Авансова вноска")]
        PaymentInAdvance = 4,

        [Description("5 - Парт. Номер на имот")]
        PartPropertyNumber = 5,

        [Description("6 - Постановление за принудително събиране")]
        CompulsoryCollectionOrder = 6,

        [Description("9 - Други")]
        Others = 9
    }
}
