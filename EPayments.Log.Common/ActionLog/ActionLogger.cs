using System;
using System.Collections.Generic;
using NLog;
using Microsoft.Owin;
using EPayments.Log.NLog;
using Newtonsoft.Json;
using EPayments.Model.Enums;

namespace EPayments.Log.ActionLogger
{
    internal class ActionLogger : IActionLogger
    {
        public static readonly Logger Logger = LogManager.GetLogger("dataLogger");

        private ILogger logger;
        private IOwinContext owinContext;

        public ActionLogger(
            ILogger logger,
            IOwinContext owinContext)
        {
            this.logger = logger;
            this.owinContext = owinContext;
        }

        public void LogAction(string clientId, object postData, object responseData)
        {
            try
            {
                string postJson = postData != null ? JsonConvert.SerializeObject(postData) : null;
                string responseJson = responseData != null ? JsonConvert.SerializeObject(responseData) : null;

                string action = null;

                this.LogAction(
                    action,
                    clientId,
                    postJson,
                    responseJson);
            }
            catch (Exception ex)
            {
                this.logger.Log(LogLevel.Error, ex.Message, ex);
            }
        }

        private void LogAction(string action, string clientId, string postData, string responseData)
        {
            LogEventInfo logEvent = new LogEventInfo();
            logEvent.Level = global::NLog.LogLevel.Info;
            logEvent.Properties["Action"] = action;
            logEvent.Properties["ClientId"] = clientId;
            logEvent.Properties["PostData"] = postData;
            logEvent.Properties["ResponseData"] = responseData;
            logEvent.Properties["RawUrl"] = this.owinContext.Request.Uri.PathAndQuery;
            logEvent.Properties["RequestId"] = ((NLogLogger)this.logger).requestId;
            logEvent.Properties["RemoteIpAddress"] = this.owinContext.Request.RemoteIpAddress;

            ActionLogger.Logger.Log(logEvent);
        }
    }
}
