using EPayments.Job.Host.Core;
using Polly;
using System;
using NLog;
using EPayments.Common;

namespace EPayments.Job.Host.Jobs
{
    public static class BoricaRetryPolicy
    {
        private static readonly int boricaRetryPeriodInSeconds = AppSettings.EPaymentsJobHost_BoricaRetryPeriodInSeconds;
        private static readonly int boricaRetryCount = AppSettings.EPaymentsJobHost_BoricaRetryCount;

        public static Policy GetBoricaRetryPolicy(JobName jobName)
        {
            return Policy
                 .Handle<Exception>()
                 .WaitAndRetry(boricaRetryCount,
                 sleepDurationProvider: (retryCount) =>
                 {
                     return TimeSpan.FromSeconds(boricaRetryPeriodInSeconds);
                 },
                 onRetry: (exception, timespan, retryCount, context) =>
                 {
                     var logMessage = $"Грешка при опит #{retryCount} за взимане на данни от Борика. Грешка: {exception.Message}.";

                     JobLogger
                        .Get(jobName)
                        .Log(LogLevel.Info, logMessage);
                 });
        }
    }
}