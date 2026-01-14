using Autofac;
using Microsoft.Owin;
using Owin;
using Autofac.Integration.Mvc;
using EPayments.Model;
using EPayments.Data;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.Repositories.Implementations;
using EPayments.EDelivery;
using EPayments.Distributions.Implementations;
using EPayments.Distributions.Interfaces;
using EPayments.Admin.Auth;
using EPayments.CVPosTransaction.Manager;

[assembly: OwinStartupAttribute(typeof(EPayments.Admin.Startup))]
namespace EPayments.Admin
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
            builder.RegisterModule(new EDeliveryRegisterModule());
            builder.RegisterFilterProvider();

            builder.RegisterType<DistributionRepository>().As<IDistributionRepository>();
            builder.RegisterType<DistributionFactory>().As<IDistributionFactory>();
            builder.RegisterType<PaymentRequestRepository>().As<IPaymentRequestRepository>();
            builder.RegisterType<EquationControlsRepository>().As<IEquationControlsRepository>();
            builder.RegisterType<CVPosRegisterManager>().As<ICVPosRegisterManager>();

            builder.RegisterFilterProvider();
            builder.RegisterType<AdminAuthorizeAttribute>().PropertiesAutowired();

            return builder.Build();
        }
    }
}
