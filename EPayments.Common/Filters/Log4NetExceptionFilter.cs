using System.Web.Mvc;
using log4net;

namespace EPayments.Common.Filters
{
    public class Log4NetExceptionFilter : HandleErrorAttribute
    {
        private static readonly ILog Logger = LogManager.GetLogger(nameof(Log4NetExceptionFilter));

        public override void OnException(ExceptionContext filterContext)
        {
            Logger.Error(filterContext.Exception);

            base.OnException(filterContext);
        }
    }
}
