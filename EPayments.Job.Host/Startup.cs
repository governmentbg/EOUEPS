using System;
using System.Collections.Generic;
using System.IO;
using Autofac;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;
using Owin;
using System.Threading;
using EPayments.Job.Host.Core;
using Microsoft.Owin;
using AppFunc = System.Func<System.Collections.Generic.IDictionary<string, object>, System.Threading.Tasks.Task>;
using EPayments.Job.Host.Email;
using EPayments.Model;
using EPayments.Data;
using EPayments.Common.Owin;
using EPayments.Common;
using EPayments.Job.Host.EserviceNotification;
using EPayments.Job.Host.ExpiredRequest;
using EPayments.Common.Helpers;
using EPayments.EventRegister;
using EPayments.Job.Host.EventRegisterNotification;
using EPayments.Distributions;
using EPayments.EDelivery;
using EPayments.Job.Host.Jobs.EDeliveryNotification;
using EPayments.CVPosTransaction;
using EPayments.Job.Host.Jobs.CVPosTransaction;
using EPayments.Job.Host.Jobs.BoricaUnprocessedRequests;
using EPayments.Job.Host.Jobs.Distributions;
using EPayments.Job.Host.Jobs.CVPosTransactionFix;

namespace EPayments.Job.Host
{
    public class Startup
    {
        public static JsonSerializerSettings JsonSerializerSettings = new JsonSerializerSettings()
        {
#if DEBUG
            Formatting = Formatting.Indented,
#else
            Formatting = Formatting.None,
#endif
            ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
            NullValueHandling = NullValueHandling.Include,
            DefaultValueHandling = DefaultValueHandling.Include,
            DateFormatHandling = DateFormatHandling.IsoDateFormat,
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            Converters = new List<JsonConverter>()
            {
                new StringEnumConverter { CamelCaseText = true }
            }
        };

        public void Configuration(IAppBuilder app)
        {
            var container = CreateAutofacContainer();

            Configure(app, container);
        }

        public static IContainer CreateAutofacContainer()
        {
            var builder = new ContainerBuilder();
            builder.RegisterModule(new ModelModule());
            builder.RegisterModule(new DataModule());
            builder.RegisterModule(new EventRegisterModule());
            builder.RegisterModule(new DistributionModule());
            builder.RegisterModule(new EDeliveryRegisterModule());
            builder.RegisterModule(new CVPosRegisterModule());

            builder.RegisterType<JobHost>();

            if (AppSettings.EPaymentsJobHost_EmailJobEnabled)
            {
                builder.RegisterType<EmailJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_EserviceNotificationJobEnabled)
            {
                builder.RegisterType<EserviceNotificationJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_EventRegisterNotificationJobEnabled)
            {
                builder.RegisterType<EventRegisterNotificationJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_ExpiredRequestJobEnabled)
            {
                builder.RegisterType<ExpiredRequestJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_ProcessTransactionFilesJobEnabled)
            {
                builder.RegisterType<ProcessTransactionFilesJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_UnprocessedVposRequestsJobEnabled)
            {
                builder.RegisterType<UnprocessedVposRequestsJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_EDeliveryNotificationJobEnabled)
            {
                builder.RegisterType<EDeliveryNotificationJob>().As<IJob>();
            }

            if (AppSettings.EPaymentsJobHost_DistributionJobEnabled)
            {
                builder.RegisterType<DistributionJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_CVPosTransactionJobEnabled)
            {
                builder.RegisterType<CVPosTransactionJob>().As<IJob>();
            }
            if (AppSettings.EPaymentsJobHost_CVPosTransactionFixJobEnabled)
            {
                builder.RegisterType<CVPosTransactionFixJob>().As<IJob>();
            }

            if (AppSettings.EPaymentsJobHost_BoricaUnprocessedRequestsEnabled)
            {
                builder.RegisterType<BoricaUnprocessedRequestsJob>().As<IJob>();
            }

            return builder.Build();
        }

        public static void Configure(IAppBuilder app, IContainer container)
        {
            //return 200 OK / "Running" on all requests
            app.Use(new Func<AppFunc, AppFunc>(next => async (env) =>
            {
                env["owin.ResponseStatusCode"] = 200;

                ((IDictionary<string, string[]>)env["owin.ResponseHeaders"])["Content-Type"] = new[] { "text/plain" };

                using (StreamWriter sw = new StreamWriter(((Stream)env["owin.ResponseBody"])))
                {
                    await sw.WriteAsync(String.Format("Running... (revision #{0})", GitRevisionHelper.GetRevision()));
                }
            }));

            StartJobs(container, app);
            app.DisposeContainerOnShutdown(container);
        }

        public static void StartJobs(IContainer container, IAppBuilder app)
        {
            var context = new OwinContext(app.Properties);
            var token = context.Get<CancellationToken>("host.OnAppDisposing");

            var jobs = container.Resolve<IJob[]>();

            foreach (var job in jobs)
            {
                container.Resolve<JobHost>().Start(job, token);
            }
        }
    }
}
