using Autofac;
using Microsoft.Owin;
using Owin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPayments.Common.Owin
{
    public static class AppBuilderExtensions
    {
        public static void DisposeContainerOnShutdown(this IAppBuilder app, IContainer container)
        {
            var context = new OwinContext(app.Properties);
            var token = context.Get<CancellationToken>("host.OnAppDisposing");

            if (token != CancellationToken.None)
            {
                token.Register(container.Dispose);
            }
        }
    }
}
