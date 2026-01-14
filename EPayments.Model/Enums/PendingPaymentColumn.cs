using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum PendingPaymentColumn
    {
        PaymentId,
        CreateDate,
        ExpirationDate,
        ServiceProvider,
        PaymentReason,
        Amount
    }
}
