using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Job.Host.Core
{
    public enum JobName
    {
        Email,
        EserviceNotification,
        EventRegisterNotification,
        ExpiredRequest,
        ProcessTransactionFiles,
        UnprocessedVposRequests,
        DistributionJob,
        EDeliveryNotification,
        CVPosTransaction,
        CVPosTransactionFix,
        BoricaUnprocessedRequestsJob
    }

    public class JobLogger
    {
        private Logger logger { get; set; }

        private JobLogger() { }

        public void Log(LogLevel logLevel, string message)
        {
            this.logger.Log(CreateLogEventInfo(logLevel, message, null));
        }

        public void Log(LogLevel logLevel, string message, Exception exception)
        {
            this.logger.Log(CreateLogEventInfo(logLevel, message, exception));
        }

        private LogEventInfo CreateLogEventInfo(
            LogLevel level,
            string message,
            Exception exception)
        {
            var logEvent = new LogEventInfo();

            logEvent.Level = level;
            logEvent.Message = message;
            logEvent.Exception = exception;

            return logEvent;
        }

        private static Dictionary<string, JobLogger> LoggersDictionary { get; set; }

        public static JobLogger Get(JobName job)
        {
            if (LoggersDictionary == null)
            {
                LoggersDictionary = new Dictionary<string, JobLogger>();
            }

            string loggerName = null;
            switch (job)
            {
                case JobName.Email:
                    loggerName = "emailJobLogger";
                    break;
                case JobName.EserviceNotification:
                    loggerName = "eserviceNotificationJobLogger";
                    break;
                case JobName.EventRegisterNotification:
                    loggerName = "eventRegisterNotificationJobLogger";
                    break;
                case JobName.ExpiredRequest:
                    loggerName = "expiredRequestJobLogger";
                    break;
                case JobName.ProcessTransactionFiles:
                    loggerName = "processTransactionFilesJobLogger";
                    break;
                case JobName.UnprocessedVposRequests:
                    loggerName = "unprocessedVposRequestsJobLogger";
                    break;
                case JobName.DistributionJob:
                    loggerName = "distributionJobLogger";
                    break;
                case JobName.EDeliveryNotification:
                    loggerName = "eDeliveryNotificationJobLogger";
                    break;
                case JobName.CVPosTransaction:
                    loggerName = "cVPosTransactionJobLogger";
                    break;
                case JobName.BoricaUnprocessedRequestsJob:
                    loggerName = "boricaUnprocessedRequestsJobLogger";
                    break;
                case JobName.CVPosTransactionFix:
                    loggerName = "cVPosTransactionFixJobLogger";
                    break;
                default:
                    throw new ArgumentException();
            }

            if (!LoggersDictionary.ContainsKey(loggerName))
            {
                LoggersDictionary.Add(loggerName,
                    new JobLogger
                    {
                        logger = LogManager.GetLogger(loggerName)
                    });
            }

            return LoggersDictionary[loggerName];
        }
    }
}