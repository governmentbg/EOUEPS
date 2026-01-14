using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace EPayments.Job.Host.Core
{
    public interface IJob : IDisposable
    {
        string Name { get; }
        TimeSpan Period { get; }
        void Action(CancellationToken token);
    }
}
