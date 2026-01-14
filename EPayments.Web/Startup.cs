using Autofac;
using Microsoft.Owin;
using Owin;
using Autofac.Integration.Mvc;
using EPayments.Model;
using EPayments.Data;
using EPayments.CertificateUtils;

[assembly: OwinStartupAttribute(typeof(EPayments.Web.Startup))]
namespace EPayments.Web
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }

        public static IContainer CreateAutofacContainer()
        {
            var builder = new ContainerBuilder();

            builder.RegisterControllers(typeof(MvcApplication).Assembly);

            builder.RegisterModule(new ModelModule());
            builder.RegisterModule(new DataModule());
            builder.RegisterModule(new CertificateUtilsModule());
            builder.RegisterFilterProvider();

            return builder.Build();
        }
    }
}
