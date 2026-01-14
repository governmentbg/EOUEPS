using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum NotificationStatus
    {
        [Description("Чака за изпращане")]
        Pending = 1,

        [Description("Изпратен")]
        Sent = 2,

        [Description("Грешка")]
        Error = 3,

        [Description("Прекратено поради невъзможност за изпращане")]
        Terminated = 4,
    }
}
