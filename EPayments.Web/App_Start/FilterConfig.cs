using EPayments.Common.Filters;
using System.Web.Mvc;

namespace EPayments.Web
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            filters.Add(new Log4NetTraceFilter());
            filters.Add(new Log4NetExceptionFilter());
        }
    }
}
