using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autofac;
using Autofac.Features.Indexed;
using Autofac.Integration.Owin;
using EPayments.Log.Owin;
using Microsoft.Owin;
using NLog;

namespace EPayments.Log.NLog
{
    public class NLogLogger : ILogger
    {
        public static readonly string ServiceLoggerName = "EPayments.Service.Host";
        private static readonly Logger InfoLogger = LogManager.GetLogger("infoLogger");

        internal string appName;
        private IDictionary<string, Func<IOwinContext, string>> customProperties;
        private IOwinContext owinContext;
        internal string requestId;

        public NLogLogger(string appName)
        {
            this.appName = appName;
        }

        public NLogLogger(string appName, IOwinContext owinContext, IDictionary<string, Func<IOwinContext, string>> customProperties)
        {
            this.appName = appName;
            this.owinContext = owinContext;
            this.requestId = Guid.NewGuid().ToString();
            this.customProperties = customProperties ?? new Dictionary<string, Func<IOwinContext, string>>();
        }

        public void Log(LogLevel logLevel, string message)
        {
            this.Log(logLevel, message, null);
        }

        public void Log(LogLevel logLevel, string message, Exception exception)
        {
            InfoLogger.Log(this.CreateLogEventInfo(logLevel, message, exception));
        }

        //this method should be thread safe and should not modify the logger!!!
        private LogEventInfo CreateLogEventInfo(
            LogLevel level,
            string message,
            Exception exception)
        {
            global::NLog.LogLevel nlogLevel;
            switch (level)
            {
                case LogLevel.Off:
                    nlogLevel = global::NLog.LogLevel.Off;
                    break;
                case LogLevel.Fatal:
                    nlogLevel = global::NLog.LogLevel.Fatal;
                    break;
                case LogLevel.Error:
                    nlogLevel = global::NLog.LogLevel.Error;
                    break;
                case LogLevel.Warn:
                    nlogLevel = global::NLog.LogLevel.Warn;
                    break;
                case LogLevel.Info:
                    nlogLevel = global::NLog.LogLevel.Info;
                    break;
                case LogLevel.Debug:
                    nlogLevel = global::NLog.LogLevel.Debug;
                    break;
                case LogLevel.Trace:
                    nlogLevel = global::NLog.LogLevel.Trace;
                    break;
                default:
                    throw new Exception("Unknown log level");
            }

            var logEvent = new LogEventInfo();

            logEvent.Level = nlogLevel;
            logEvent.Message = message;
            logEvent.Exception = exception;
            logEvent.Properties["AppName"] = this.appName;

            if (this.owinContext != null)
            {
                logEvent.Properties["RemoteIpAddress"] =
                    this.appName == NLogLogger.ServiceLoggerName && this.owinContext.Request.Headers.ContainsKey("X-Original-Client-IP") ?
                    this.owinContext.Request.Headers["X-Original-Client-IP"] :
                    this.owinContext.Request.RemoteIpAddress;
                logEvent.Properties["Method"] = this.owinContext.Request.Method;
                logEvent.Properties["PathAndQuery"] = this.owinContext.Request.Uri.PathAndQuery;
                logEvent.Properties["UserAgent"] = this.owinContext.Request.Headers["User-Agent"];
                logEvent.Properties["RequestId"] = this.requestId;

                double? elapsedMilliseconds = null;
                long? startTicks = this.owinContext.Get<long?>(LoggingMiddleware.TimerTicksAtEntryOwinEnvKey);
                if (startTicks.HasValue)
                {
                    double millisecondsPerTick = 1000d / Stopwatch.Frequency;
                    elapsedMilliseconds = (long)Math.Floor((Stopwatch.GetTimestamp() - startTicks.Value) * millisecondsPerTick);
                }

                logEvent.Properties["ElapsedMilliseconds"] = elapsedMilliseconds;

                string status;
                if (this.owinContext.Request.CallCancelled.IsCancellationRequested)
                {
                    status = "Cancelled";
                }
                else
                {
                    status = string.Format("{0}: {1}", this.owinContext.Response.StatusCode, this.owinContext.Response.ReasonPhrase);
                }

                logEvent.Properties["Status"] = status;

                foreach (var prop in this.customProperties)
                {
                    logEvent.Properties[prop.Key] = prop.Value(this.owinContext);
                }
            }

            return logEvent;
        }
    }
}
