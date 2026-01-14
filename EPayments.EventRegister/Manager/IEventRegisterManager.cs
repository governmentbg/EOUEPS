using EPayments.EventRegister.DataObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.EventRegister.Manager
{
    public interface IEventRegisterManager
    {
        bool IsReady { get; }
        void LogEvent(EventDO eventDO);
        Task LogEventAsync(EventDO eventDO);
    }
}
