using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http.Controllers;
using log4net;
using Microsoft.Owin;
using Newtonsoft.Json;
using ActionFilterAttribute = System.Web.Http.Filters.ActionFilterAttribute;

namespace EPayments.Service.Api.Common.CustomExceptions
{
    public class CustomeTraceFilter : ActionFilterAttribute
    {
        private static readonly ILog Logger = LogManager.GetLogger(nameof(CustomeTraceFilter));

        public override void OnActionExecuting(HttpActionContext filterContext)
        {
            string bodyText;
            var ip = GetClientIpAddress(filterContext.Request);
            var requestBody = filterContext.ActionArguments.Values.FirstOrDefault();
            bodyText = JsonConvert.SerializeObject(requestBody);

            Logger.Info($"Path :{filterContext.Request.RequestUri}, RequestBody:{bodyText}, ClientIP:{ip}");

            base.OnActionExecuting(filterContext);
        }

        private string GetClientIpAddress(HttpRequestMessage request)
        {
            if (request.Properties.ContainsKey("MS_HttpContext"))
            {
                return IPAddress.Parse(((HttpContextBase)request.Properties["MS_HttpContext"]).Request.UserHostAddress).ToString();
            }
            if (request.Properties.ContainsKey("MS_OwinContext"))
            {
                return IPAddress.Parse(((OwinContext)request.Properties["MS_OwinContext"]).Request.RemoteIpAddress).ToString();
            }
            return String.Empty;
        }
    }
}
