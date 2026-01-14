using EPayments.EventRegister.EventRegisterService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.EventRegister.DataObjects
{
    public class EventDO
    {
        public EventDO()
        {
            this.Time = DateTime.UtcNow;

        }
        public EventTypeEnum EventType { get; set; }

        public string EventDescription { get; set; }

        public DateTime Time { get; set; }

        public string DocumentRegNumber { get; set; }

        public string ServiceOID { get; set; }
    }
}
