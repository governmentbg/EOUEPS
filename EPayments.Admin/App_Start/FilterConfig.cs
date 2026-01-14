using System.Web.Mvc;
using EPayments.Common.Filters;

namespace EPayments.Admin
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            filters.Add(new Log4NetExceptionFilter());
        }
    }
}
