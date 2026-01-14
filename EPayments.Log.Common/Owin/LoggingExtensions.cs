using Owin;

namespace EPayments.Log.Owin
{
    public static class LoggingExtensions
    {
        public static IAppBuilder UseLogging(this IAppBuilder app, ILoggerFactory loggerFactory)
        {
            return app.Use<LoggingMiddleware>(loggerFactory);
        }
    }
}
