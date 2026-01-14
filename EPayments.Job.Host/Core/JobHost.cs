using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Hosting;
using System.Diagnostics;
using NLog;
using EPayments.Job.Host.Email;

namespace EPayments.Job.Host.Core
{
    public class JobHost : IDisposable
    {
        private readonly object jobLock = new object();

        private IJob job;
        private Timer timer;

        // reads & writes of bool are atomic, so no lock is required
        public bool IsShuttingDown { get; private set; }

        public void Start(IJob job, CancellationToken token)
        {
            if (job == null)
            {
                throw new ArgumentNullException("job");
            }

            this.job = job;
            JobLogger.Get(GetJobName(this.job.Name))
                .Log(LogLevel.Info, string.Format("Initializing {0}", this.job.Name));
            this.timer = new Timer((sender) => this.DoAction(sender, token));
            this.timer.Change(TimeSpan.FromSeconds(0), this.job.Period);
        }

        private void DoAction(object sender, CancellationToken token)
        {
            if (this.IsShuttingDown)
            {
                return;
            }

            // DoAction returns immediately if the previous action has not finished
            if (Monitor.TryEnter(this.jobLock))
            {
                try
                {
                    this.job.Action(token);
                }
                catch (Exception e)
                {
                    JobLogger.Get(GetJobName(this.job.Name))
                        .Log(LogLevel.Error, string.Format("Error while running {0}", this.job.Name), e);
                }
                finally
                {
                    Monitor.Exit(this.jobLock);
                }
            }
        }

        public void Dispose()
        {
            if (this.timer != null)
            {
                this.timer.Dispose();
                this.timer = null;
            }

            if (this.job != null)
            {
                JobName jobName = GetJobName(this.job.Name);

                this.job.Dispose();
                JobLogger.Get(jobName)
                    .Log(LogLevel.Info, string.Format("Disposed {0}", this.job.Name));
                this.job = null;
            }
        }

        private JobName GetJobName(string jobName)
        {
            switch (jobName)
            {
                case "EmailJob":
                    return JobName.Email;
                case "EserviceNotificationJob":
                    return JobName.EserviceNotification;
                case "EventRegisterNotificationJob":
                    return JobName.EventRegisterNotification;
                case "ExpiredRequestJob":
                    return JobName.ExpiredRequest;
                case "ProcessTransactionFilesJob":
                    return JobName.ProcessTransactionFiles;
                case "UnprocessedVposRequestsJob":
                    return JobName.UnprocessedVposRequests;
                case "EDeliveryNotificationJob":
                    return JobName.EDeliveryNotification;
                case "DistributionJob":
                    return JobName.DistributionJob;
                case "CVPosTransactionJob":
                    return JobName.CVPosTransaction;
                case "BoricaUnprocessedRequestsJob":
                    return JobName.BoricaUnprocessedRequestsJob;
                case "CVPosTransactionFixJob":
                    return JobName.CVPosTransactionFix;
                default:
                    throw new NotImplementedException();
            }
        }
    }
}