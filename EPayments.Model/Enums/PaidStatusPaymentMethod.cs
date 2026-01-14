using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum PaidStatusPaymentMethod
    {
        [Description("Друг")]
        Other = 1,

        [Description("На каса")]
        CashDesk = 2,
    }
}
