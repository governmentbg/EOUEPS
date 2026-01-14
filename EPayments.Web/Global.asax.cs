using Autofac;
using Autofac.Integration.Mvc;
using EPayments.Common;
using log4net;
using System;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace EPayments.Web
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            IContainer container = Startup.CreateAutofacContainer();
            DependencyResolver.SetResolver(new AutofacDependencyResolver(container));

            ModelBinders.Binders.Add(typeof(decimal), new DecimalModelBinder(null));
        }
    }
}
