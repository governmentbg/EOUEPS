using Autofac.Features.OwnedInstances;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using EPayments.Job.Host.Core;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Common.Helpers;
using System.Net;
using NLog;
using EPayments.Model.Models;

namespace EPayments.Job.Host.Email
{
    public class EmailJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory;
        private object syncRoot = new object();
        private bool disposed;
        private int batchSize;
        private TimeSpan period;
        private int maxFailedAttempts;
        private TimeSpan failedAttemptTimeout;
        private int parallelTasks;
        private string mailServer;
        private int successes;
        private int failures;

        public EmailJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;

            this.batchSize = AppSettings.EPaymentsJobHost_EmailJobBatchSize;
            this.period = TimeSpan.FromSeconds(AppSettings.EPaymentsJobHost_EmailJobPeriodInSeconds);
            this.maxFailedAttempts = AppSettings.EPaymentsJobHost_EmailJobMaxFailedAttempts;
            this.failedAttemptTimeout = TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_EmailJobFailedAttemptTimeoutInMinutes);
            this.parallelTasks = AppSettings.EPaymentsJobHost_EmailJobParallelTasks;
            this.mailServer = AppSettings.EPaymentsJobHost_EmailJobMailServer;
        }

        public string Name
        {
            get { return "EmailJob"; }
        }

        public TimeSpan Period
        {
            get { return this.period; }
        }

        public void Action(CancellationToken ct)
        {
            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var systemRepository = factory.Value.Item3;

                GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.EmailJobLastInvocationTime);
                if (lastInvocationTime != null)
                {
                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();
                }
            }

            IList<int> pendingEmailIds = new List<int>();

            try
            {
                if (disposed)
                {
                    return;
                }

                using (var factory = this.dependencyFactory())
                {
                    var unitOfWork = factory.Value.Item1;
                    var jobRepository = factory.Value.Item2;

                    pendingEmailIds = jobRepository.GetPendingEmailIds(this.batchSize, this.maxFailedAttempts, this.failedAttemptTimeout);
                }

                if (pendingEmailIds.Any())
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();

                    this.successes = 0;
                    this.failures = 0;

                    SendParallel(ct, pendingEmailIds).Wait();

                    sw.Stop();
                    JobLogger.Get(JobName.Email)
                        .Log(LogLevel.Info, string.Format("Email batch finished in {0}ms - {1} emails send, {2} failures of total {3} emails.", sw.ElapsedMilliseconds, this.successes, this.failures, pendingEmailIds.Count));
                }
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.Email)
                    .Log(LogLevel.Error, string.Format("Job was canceled due to a token cancellation request; Email batch finished with {0} emails send, {1} failures of total {2} emails.", this.successes, this.failures, pendingEmailIds.Count), ex);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.Email)
                    .Log(LogLevel.Error, ex.Message, ex);
            }
        }

        private Task SendParallel(CancellationToken ct, IList<int> pendingEmailIds)
        {
            ConcurrentQueue<int> mailIds = new ConcurrentQueue<int>(pendingEmailIds);

            int numberOfParallelTasks = Math.Min(mailIds.Count, this.parallelTasks);

            var parallelTasks = Enumerable.Range(0, numberOfParallelTasks)
                .Select(pt => Task.Run(() => Send(ct, mailIds), ct))
                .ToArray();

            return Task.WhenAll(parallelTasks);
        }

        private async Task Send(CancellationToken ct, ConcurrentQueue<int> mailIds)
        {
            int mailId;
            EPayments.Model.Models.Email email;

            using (SmtpClient smtpClient = CreateSmtpClient())
            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var jobRepository = factory.Value.Item2;

                while (mailIds.TryDequeue(out mailId))
                {
                    if (disposed)
                    {
                        break;
                    }

                    ct.ThrowIfCancellationRequested();

                    email = jobRepository.GetEmailById(mailId);

                    try
                    {
                        TemplateConfig templateConfig = TemplateConfig.Get(email.MailTemplateName);

                        MailAddress from = new MailAddress(templateConfig.Sender);
                        MailAddress to = new MailAddress(email.Recipient);

                        MailMessage mailMessage = new MailMessage(from, to);
                        mailMessage.Subject = templateConfig.MailSubject;
                        mailMessage.Body = BuildEmailBody(templateConfig.TemplateFileName, email.Context);
                        mailMessage.SubjectEncoding = System.Text.Encoding.GetEncoding(1251);
                        mailMessage.BodyEncoding = System.Text.Encoding.UTF8;
                        mailMessage.IsBodyHtml = templateConfig.IsBodyHtml;
                        mailMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.None;

                        await smtpClient.SendMailAsync(mailMessage);

                        email.SetStatus(NotificationStatus.Sent);

                        Interlocked.Increment(ref this.successes);
                    }
                    catch (SmtpException smtpEx)
                    {
                        var exception = "SmtpException: " + Enum.GetName(typeof(SmtpStatusCode), smtpEx.StatusCode);

                        email.IncrementFailedAttempts(exception);
                        if (email.FailedAttempts >= this.maxFailedAttempts)
                        {
                            email.SetStatus(NotificationStatus.Terminated);
                        }

                        JobLogger.Get(JobName.Email).Log(LogLevel.Warn, exception);
                        Interlocked.Increment(ref this.failures);
                    }
                    catch (Exception ex)
                    {
                        email.IncrementFailedAttempts(ex.Message);
                        email.SetStatus(NotificationStatus.Error);

                        JobLogger.Get(JobName.Email).Log(LogLevel.Error, ex.Message, ex);
                        Interlocked.Increment(ref this.failures);
                    }

                    unitOfWork.Save();
                }
            }
        }

        private SmtpClient CreateSmtpClient()
        {
            SmtpClient client;

            if (!AppSettings.EmailJobUseTestGmailServer)
            {
                client = new SmtpClient(this.mailServer);
            }
            else
            {
                client = new SmtpClient()
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,                    
                    UseDefaultCredentials = false,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential("testepaymentsdaeu@gmail.com", "Trudna_Parola42")
                };
            }

            return client;
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;

                JobLogger.Get(JobName.Email).Log(LogLevel.Info, "Email job disposed");
            }
        }

        private string BuildEmailBody(string templateFileName, string context)
        {
            string templatePath = GetTemplatePath(templateFileName);
            JObject templateModel = context != null ? JObject.Parse(context) : null;

            return RenderHelper.RenderHtmlByRazorTemplate(templateFileName, templatePath, templateModel);
        }

        private string GetTemplatePath(string templateName)
        {
            string assemblyPath = (new System.Uri(Assembly.GetExecutingAssembly().CodeBase)).LocalPath;
            string binPath = System.IO.Path.GetDirectoryName(assemblyPath);
            string templateFullPath = String.Format(@"{0}\Jobs\Email\Templates\{1}", binPath, templateName);

            return templateFullPath;
        }
    }
}
