using Autofac;
using EPayments.Log.Owin;
using Microsoft.Owin;

namespace EPayments.Log
{
    public class LogModule : Module
    {
        private ILoggerFactory loggerFactory;

        public LogModule(ILoggerFactory loggerFactory)
        {
            this.loggerFactory = loggerFactory;
        }

        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.Register(c =>
            {
                IOwinContext owinContext;
                if (c.TryResolve<IOwinContext>(out owinContext))
                {
                    return owinContext.Get<ILogger>(LoggingMiddleware.LoggerOwinEnvKey);
                }
                else
                {
                    return loggerFactory.Create();
                }
            });
        }
    }
}
