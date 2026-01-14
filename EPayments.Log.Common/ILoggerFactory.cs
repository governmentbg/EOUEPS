using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Owin;

namespace EPayments.Log
{
    public interface ILoggerFactory
    {
        ILogger Create();
        ILogger Create(IOwinContext owinContext);
    }
}
