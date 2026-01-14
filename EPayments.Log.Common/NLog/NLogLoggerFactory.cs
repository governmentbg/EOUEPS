using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Owin;

namespace EPayments.Log.NLog
{
    public class NLogLoggerFactory : ILoggerFactory
    {
        private string appName;
        private IDictionary<string, Func<IOwinContext, string>> customProperties;

        public NLogLoggerFactory(string appName)
        {
            this.appName = appName;
        }

        public NLogLoggerFactory(string appName, IDictionary<string, Func<IOwinContext, string>> customProperties)
        {
            this.appName = appName;
            this.customProperties = customProperties;
        }

        public ILogger Create()
        {
            return new NLogLogger(this.appName);
        }

        public ILogger Create(IOwinContext owinContext)
        {
            return new NLogLogger(this.appName, owinContext, this.customProperties);
        }
    }
}
