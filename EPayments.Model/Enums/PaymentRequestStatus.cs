using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum PaymentRequestStatus
    {
        [Description("Очаква плащане")]
        Pending = 1,

        [Description("Платено с карта")]
        Authorized = 2,

        [Description("Платено по банков път")]
        Ordered = 3,

        [Description("Получено плащане")]
        Paid = 4,

        [Description("Изтекъл срок")]
        Expired = 5,

        [Description("Отказано")]
        Canceled = 6,

        [Description("Прекратена услуга")]
        Suspended = 7,

        [Description("В процес на обработка")]
        InProcess = 9
    }

    public static class PaymentRequestStatusExtensions
    {
        public static string ToPaymentMethod(this PaymentRequestStatus status, bool isVposAuthorized)
        {
            string paymentMethod = String.Empty;

            switch(status)
            {
                case PaymentRequestStatus.Authorized:
                    paymentMethod = "Карта";
                    break;
                case PaymentRequestStatus.Ordered:
                    paymentMethod = "По банков път";
                    break;
                case PaymentRequestStatus.Paid:
                    paymentMethod = isVposAuthorized ? "Карта" : "По банков път";
                    break;
            }

            return paymentMethod;
        }
    }
}
