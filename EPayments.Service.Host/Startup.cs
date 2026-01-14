using System.Web.Http;
using Autofac;
using Autofac.Integration.WebApi;
using Owin;
using EPayments.Service.Api;
using Microsoft.Owin;
using EPayments.Data;
using EPayments.Documents;
using EPayments.Model;
using EPayments.Service.Api.Common.CustomExceptions;
using Microsoft.Owin.StaticFiles;
using Microsoft.Owin.FileSystems;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json.Converters;
using System.Collections.Generic;
using EPayments.Log;
using EPayments.Log.NLog;
using EPayments.Log.Owin;
using EPayments.Log.Api;
using System.Web.Http.ExceptionHandling;
using Swashbuckle.Application;
using EPayments.Common.Filters;

namespace EPayments.Service.Host
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

        public static ILoggerFactory LoggerFactory = new NLogLoggerFactory(NLogLogger.ServiceLoggerName);

        public void Configuration(IAppBuilder app)
        {
            var container = CreateAutofacContainer();

            Configure(app, container);
        }

        public static IContainer CreateAutofacContainer()
        {
            var builder = new ContainerBuilder();
            builder.RegisterModule(new DataModule());
            builder.RegisterModule(new ModelModule());
            builder.RegisterModule(new DocumentsModule());
            builder.RegisterModule(new ServiceApiModule());
            builder.RegisterModule(new LogModule(LoggerFactory));
            return builder.Build();
        }

        public static void Configure(IAppBuilder app, IContainer container)
        {
            app.UseAutofacMiddleware(container);
            app.UseLogging(LoggerFactory);
            ConfigureWebApi(app, container);
            ConfigureStaticFiles(app);
        }

        public static void ConfigureStaticFiles(IAppBuilder app)
        {
            app.UseDirectoryBrowser(new DirectoryBrowserOptions
            {
                RequestPath = new PathString("/documents"),
                FileSystem = new PhysicalFileSystem("./Documents")
            });

            app.UseStaticFiles(new StaticFileOptions
            {
                RequestPath = new PathString("/documents"),
                FileSystem = new PhysicalFileSystem("./Documents"),
                ServeUnknownFileTypes = true
            });
        }

        public static void ConfigureWebApi(IAppBuilder app, IContainer container)
        {
            HttpConfiguration config = new HttpConfiguration();

            ContainerBuilder builder = new ContainerBuilder();
            builder.RegisterHttpRequestMessage(config);
            builder.Update(container);

            //auth
            config.Filters.Add(new CustomExceptionFilter());
            config.Filters.Add(new CustomeTraceFilter());

            //json serialization
            config.Formatters.JsonFormatter.SerializerSettings = JsonSerializerSettings;
            JsonConvert.DefaultSettings = () => JsonSerializerSettings;

            //logging
            config.Services.Add(typeof(IExceptionLogger), new LoggingMiddlewareExceptionLogger(null));

            //routing
            config.MapHttpAttributeRoutes();

            //disable showing of detailed errors
            config.IncludeErrorDetailPolicy = IncludeErrorDetailPolicy.Never;

            app.UseAutofacWebApi(config);
            app.UseWebApi(config);

            config
             .EnableSwagger(c => c.SingleApiVersion("v1", "EPayments.Service.Host"))
             .EnableSwaggerUi();
        }
    }
}
