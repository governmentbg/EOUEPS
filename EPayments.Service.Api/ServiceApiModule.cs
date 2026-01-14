using Autofac;
using EPayments.Service.Api.Common.XsdValidator;
using EPayments.Service.Api.Controllers.v1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api
{
    public class ServiceApiModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            //Controllers
            moduleBuilder.RegisterType<EserviceController>().InstancePerLifetimeScope();
            moduleBuilder.RegisterType<EbankingController>().InstancePerLifetimeScope();

            //Others
            moduleBuilder.RegisterType<SchemaValidator>().As<ISchemaValidator>().InstancePerLifetimeScope(); 
        }
    }
}
