using Autofac;
using Autofac.Integration.WebApi;
using Newtonsoft.Json;
using System.Net.Http;
using System.Web.Http.Filters;
using System.Collections.Generic;
using System;
using NLog;
using EPayments.Log.NLog;
using Microsoft.Owin;
using EPayments.Model.Enums;

namespace EPayments.Log.ActionLogger
{
    public class ActionLogAttribute : ActionFilterAttribute
    {
        public string ClientIdParam { get; set; }

        public ActionLogAttribute(string clientIdParam = "requestDO.ClientId")
        {
            this.ClientIdParam = clientIdParam;
        }

        public override void OnActionExecuted(HttpActionExecutedContext actionExecutedContext)
        {
            var dependencyScope = actionExecutedContext.Request.GetDependencyScope();
            var lifetimeScope = dependencyScope.GetRequestLifetimeScope();
            var logger = ((NLogLogger)lifetimeScope.Resolve<ILogger>());

            try
            {
                var owinContext = lifetimeScope.Resolve<IOwinContext>();

                string postData = JsonConvert.SerializeObject(actionExecutedContext.ActionContext.ActionArguments);

                string responseData = null;
                if (actionExecutedContext.Response != null && actionExecutedContext.Response.Content != null)
                {
                    object returnValue =
                        actionExecutedContext.Response.Content.GetType().GetProperty("Value")
                        .GetValue(actionExecutedContext.Response.Content);

                    responseData = JsonConvert.SerializeObject(returnValue);
                }

                string rawUrl = actionExecutedContext.Request.RequestUri.PathAndQuery;
                string method = actionExecutedContext.Request.RequestUri.AbsolutePath;
                string clientId = this.ClientIdParam != null ? GetClientIdValue(actionExecutedContext.ActionContext.ActionArguments, this.ClientIdParam) : null;

                LogAction(
                    method,
                    clientId,
                    postData,
                    responseData,
                    rawUrl,
                    logger.requestId,
                    owinContext.Request.RemoteIpAddress);
            }
            catch (Exception ex)
            {
                logger.Log(LogLevel.Error, ex.Message, ex);
            }
            finally
            {
                base.OnActionExecuted(actionExecutedContext);
            }
        }

        private void LogAction(string method, string clientId, string postData, string responseData, string rawUrl, string requestId, string remoteIpAddress)
        {
            LogEventInfo logEvent = new LogEventInfo();
            logEvent.Level = global::NLog.LogLevel.Info;
            logEvent.Properties["Method"] = method;
            logEvent.Properties["ClientId"] = clientId;
            logEvent.Properties["PostData"] = postData;
            logEvent.Properties["ResponseData"] = responseData;
            logEvent.Properties["RawUrl"] = rawUrl;
            logEvent.Properties["RequestId"] = requestId;
            logEvent.Properties["RemoteIpAddress"] = remoteIpAddress;

            ActionLogger.Logger.Log(logEvent);
        }

        private string GetClientIdValue(Dictionary<string, object> argumentsDictionary, string idParam)
        {
            string value = null;

            if (!String.IsNullOrWhiteSpace(idParam))
            {
                string[] keys = idParam.Split(new char[] { '.' });

                if (argumentsDictionary.ContainsKey(keys[0]))
                {
                    object propertyObj = argumentsDictionary[keys[0]];

                    if (keys.Length > 1)
                    {
                        for (int i = 1; i < keys.Length; i++)
                        {
                            var property = propertyObj.GetType().GetProperty(keys[i]);
                            if (property != null)
                            {
                                propertyObj = propertyObj.GetType().GetProperty(keys[i]).GetValue(propertyObj);
                            }
                            else
                            {
                                propertyObj = null;
                                break;
                            }
                        }
                    }

                    value = (string)propertyObj;
                }
            }

            return value;
        }
    }
}
