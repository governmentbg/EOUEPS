using Autofac;
using EPayments.Common.Data;
using EPayments.Model;

namespace EPayments.Model
{
    public class ModelModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<ModelDbConfiguration>().As<IDbConfiguration>().SingleInstance();
        }
    }
}