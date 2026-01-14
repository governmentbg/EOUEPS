using Autofac;
using EPayments.CVPosTransaction.Manager;

namespace EPayments.CVPosTransaction
{
     public class CVPosRegisterModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<CVPosRegisterManager>().As<ICVPosRegisterManager>();
        }
    }
}
