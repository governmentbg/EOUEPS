using Autofac;
using EPayments.Common;
using EPayments.EDelivery.Manager;

namespace EPayments.EDelivery
{
    public class EDeliveryRegisterModule: Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            bool useProduction = AppSettings.EPaymentsEDelivery_PrivateDeliveryUseProduction;

            if (!useProduction)
            {
                moduleBuilder.RegisterType<DeliveryRegisterManager>().As<IDeliveryRegisterManager>();
            }
            else
            {
                moduleBuilder.RegisterType<DeliveryRegisterProductionManager>().As<IDeliveryRegisterManager>();
            }
        }
    }
}
