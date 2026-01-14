using Autofac;
using EPayments.Documents.Serializer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Documents
{
    public class DocumentsModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<DocumentSerializer>().As<IDocumentSerializer>().InstancePerLifetimeScope();
        }
    }
}
