using Autofac;
using EPayments.EventRegister.Manager;

namespace EPayments.EventRegister
{
    public class EventRegisterModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<EventRegisterManager>().As<IEventRegisterManager>();
        }
    }
}
