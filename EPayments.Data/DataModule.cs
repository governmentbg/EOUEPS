using Autofac;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Data.Repositories.Implementations;
using EPayments.Data.Repositories.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data
{
    public class DataModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<UnitOfWork>().As<IUnitOfWork>().InstancePerLifetimeScope();

            //Repositories
            moduleBuilder.RegisterType<WebRepository>().As<IWebRepository>().InstancePerLifetimeScope();
            moduleBuilder.RegisterType<ApiRepository>().As<IApiRepository>().InstancePerLifetimeScope();
            moduleBuilder.RegisterType<CommonRepository>().As<ICommonRepository>().InstancePerLifetimeScope();
            moduleBuilder.RegisterType<SystemRepository>().As<ISystemRepository>().InstancePerLifetimeScope();
            moduleBuilder.RegisterType<JobRepository>().As<IJobRepository>().InstancePerLifetimeScope();
            moduleBuilder.RegisterType<PaymentRequestRepository>().As<IPaymentRequestRepository>().InstancePerLifetimeScope();
            //DisposableTuple
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,,,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,,,,,,>)).AsSelf();
            moduleBuilder.RegisterGeneric(typeof(DisposableTuple<,,,,,,,,,,>)).AsSelf();
        }
    }
}
