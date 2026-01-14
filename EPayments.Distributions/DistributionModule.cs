using Autofac;
using EPayments.Distributions.Implementations;
using EPayments.Distributions.Interfaces;

namespace EPayments.Distributions
{
    public class DistributionModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<DistributionFactory>().As<IDistributionFactory>().InstancePerDependency();
        }
    }
}
