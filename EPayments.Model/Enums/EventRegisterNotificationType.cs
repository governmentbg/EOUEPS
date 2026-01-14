using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum EventRegisterNotificationType
    {
        VposPaymentAuthorized = 1,
        PaymentRequestRegistered = 2,
        PaymentRequestDenied = 3,
    }
}
