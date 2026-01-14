using System;
using System.Linq;
using System.Net.Http;
using System.Web.Http.ExceptionHandling;
using EPayments.Log.Owin;

namespace EPayments.Log.Api
{
    public class LoggingMiddlewareExceptionLogger : ExceptionLogger
    {
        private Type[] knownExceptionTypes;

        public LoggingMiddlewareExceptionLogger()
        {
            this.knownExceptionTypes = new Type[0];
        }

        public LoggingMiddlewareExceptionLogger(Type[] knownExceptionTypes)
        {
            this.knownExceptionTypes = knownExceptionTypes;
        }

        public override void Log(ExceptionLoggerContext context)
        {
            if (knownExceptionTypes == null || !knownExceptionTypes.Contains(context.Exception.GetType()))
            {
                context.Request.GetOwinContext()
                    .Get<ILogger>(LoggingMiddleware.LoggerOwinEnvKey)
                    .Log(LogLevel.Error, context.Exception.Message, context.Exception);
            }
        }
    }
}
