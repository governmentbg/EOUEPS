using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;
using log4net;
using Microsoft.Owin;

namespace EPayments.Common.Filters
{
    public class Log4NetTraceFilter : ActionFilterAttribute
    {
        private static readonly ILog Logger = LogManager.GetLogger(nameof(Log4NetTraceFilter));

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var ip = filterContext.HttpContext.Request.UserHostAddress;
            var hostName = filterContext.HttpContext.Request.UserHostName;
            //log here requests
            var bodyStream = new StreamReader(filterContext.HttpContext.Request.InputStream);
            bodyStream.BaseStream.Seek(0, SeekOrigin.Begin);
            var bodyText = bodyStream.ReadToEnd();

            Logger.Info($"Path :{filterContext.HttpContext.Request.Path}, RequestBody:{bodyText}, QueryString:{filterContext.HttpContext.Request.RawUrl}, ClientIP:{ip}, ClientHostName:{hostName}");

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
