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
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using EPayments.Model.Models;

namespace EPayments.Job.Host.EserviceNotification
{
    public class EserviceNotificationJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory;
        private object syncRoot = new object();
        private bool disposed;
        private int batchSize;
        private TimeSpan period;
        private int parallelTasks;
        private int successes;
        private int failures;

        public EserviceNotificationJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;

            this.batchSize = AppSettings.EPaymentsJobHost_EserviceNotificationJobBatchSize;
            this.period = TimeSpan.FromSeconds(AppSettings.EPaymentsJobHost_EserviceNotificationJobPeriodInSeconds);
            this.parallelTasks = AppSettings.EPaymentsJobHost_EserviceNotificationJobParallelTasks;
        }

        public string Name
        {
            get { return "EserviceNotificationJob"; }
        }

        public TimeSpan Period
        {
            get { return this.period; }
        }

        public void Action(CancellationToken ct)
        {
            JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Info, $".net version is {Environment.Version}");

            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var systemRepository = factory.Value.Item3;

                GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.EserviceNotificationJobLastInvocationTime);
                if (lastInvocationTime != null)
                {
                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();
                }
            }

            IList<int> pendingNotifications = new List<int>();

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

                    pendingNotifications = jobRepository.GetPendingEserviceNotificationIds(this.batchSize);
                }

                if (pendingNotifications.Any())
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();

                    this.successes = 0;
                    this.failures = 0;

                    SendParallel(ct, pendingNotifications).Wait();

                    sw.Stop();
                    JobLogger.Get(JobName.EserviceNotification)
                        .Log(LogLevel.Info, string.Format("EserviceNotification batch finished in {0}ms - {1} notifications send, {2} failures of total {3} notifications.", sw.ElapsedMilliseconds, this.successes, this.failures, pendingNotifications.Count));
                }
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.EserviceNotification)
                    .Log(LogLevel.Error, string.Format("Job was canceled due to a token cancellation request; Notifications batch finished with {0} notifications send, {1} failures of total {2} notifications.", this.successes, this.failures, pendingNotifications.Count), ex);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.EserviceNotification)
                    .Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
            }
        }

        private Task SendParallel(CancellationToken ct, IList<int> pendingNotificationIds)
        {
            ConcurrentQueue<int> notificationIds = new ConcurrentQueue<int>(pendingNotificationIds);

            int numberOfParallelTasks = Math.Min(notificationIds.Count, this.parallelTasks);

            var parallelTasks = Enumerable.Range(0, numberOfParallelTasks)
                .Select(pt => Task.Run(() => Send(ct, notificationIds), ct))
                .ToArray();

            return Task.WhenAll(parallelTasks);
        }

        private async Task Send(CancellationToken ct, ConcurrentQueue<int> notificationIds)
        {
            int notificationId;
            EPayments.Model.Models.EserviceNotification notification;

            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var jobRepository = factory.Value.Item2;

                while (notificationIds.TryDequeue(out notificationId))
                {
                    try
                    {
                        if (disposed)
                        {
                            break;
                        }

                        ct.ThrowIfCancellationRequested();

                        notification = jobRepository.GetEserviceNotificationById(notificationId);
                        
                        // notifications are stopped for payment requests with algorithm 1 and 2 and request statuses different than Paid
                        int algorithmId = notification.PaymentRequest.ObligationType.AlgorithmId;
                        if ((algorithmId == 1 || algorithmId == 2) && notification.PaymentRequest.PaymentRequestStatusId != PaymentRequestStatus.Paid)
                        {
                            jobRepository.RemoveEserviceNotification(notification);
                            continue;
                        }

                        if (notification.EserviceClientId == 1611 
                            || notification.EserviceClientId == 1612 
                            || notification.EserviceClientId == 1673)
                        {
                            continue;
                        }

                        try
                        {
                            string postMediaType = "application/x-www-form-urlencoded";
                            string acceptHeaderMediaType1 = "application/json";
                            string acceptHeaderMediaType2 = "text/plain";

                            var jsonDataBytes = Encoding.UTF8.GetBytes(notification.PostData);

                            var base64Data = Convert.ToBase64String(jsonDataBytes);

                            var hmac = HmacRequestHelper.CalculateHmac(notification.EserviceClient.SecretKey, base64Data);

                            Uri notificationUri = new Uri(notification.Url);

                            string baseAddress = String.Format("{0}://{1}:{2}", notificationUri.Scheme, notificationUri.Host, notificationUri.Port);
                            string postUri = notificationUri.PathAndQuery;

                            using (var client = new HttpClient())
                            {
                                client.BaseAddress = new Uri(baseAddress);
                                client.DefaultRequestHeaders.Accept.Clear();
                                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(acceptHeaderMediaType1));
                                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(acceptHeaderMediaType2));

                                var encodedHttpBody = String.Format("ClientId={0}&Hmac={1}&Data={2}", HttpUtility.UrlEncode(notification.EserviceClient.ClientId), HttpUtility.UrlEncode(hmac), HttpUtility.UrlEncode(base64Data));
                                string encodedHttpBodyMateus = String.Format("ClientId={0}&Data={1}", HttpUtility.UrlEncode(notification.EserviceClient.ClientId), HttpUtility.UrlEncode(base64Data));

                                ServicePointManager.ServerCertificateValidationCallback += delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                                {
                                    return true;
                                };

                                try
                                {
                                    var task = client.PostAsync(postUri, new StringContent(algorithmId == 1 ? encodedHttpBodyMateus : encodedHttpBody, Encoding.UTF8, postMediaType));

                                    await task;

                                    string resultContent;
                                    if (IsNotificationAccepted(task.Result, out resultContent))
                                    {
                                        notification.SetStatus(NotificationStatus.Sent);
                                        Interlocked.Increment(ref this.successes);
                                    }
                                    else
                                    {
                                        string errorMessage = String.Format("Invalid response content. Status code - {0}, ResultContent - {1}", task.Result.StatusCode.ToString(), resultContent);
                                        throw new Exception(errorMessage);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    var exception = "HttpClientException: " + (ex.Message ?? String.Empty) + (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message ?? String.Empty : String.Empty);

                                    notification.IncrementFailedAttempts(exception);
                                    DateTime? sendNotBefore = CalculateNextSendingAttemptTime(notification.FailedAttempts);
                                    notification.SetNextSendingAttemptTime(sendNotBefore, !sendNotBefore.HasValue);

                                    JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Warn, exception);
                                    Interlocked.Increment(ref this.failures);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            notification.SetStatus(NotificationStatus.Error);
                            notification.IncrementFailedAttempts($"{ex.Message}, StackTrace -> {ex.StackTrace}");

                            JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
                            Interlocked.Increment(ref this.failures);
                        }

                        unitOfWork.Save();
                    }
                    catch (Exception ex)
                    {
                        var exception = "SendNotificationException: " + (ex.Message ?? String.Empty) + (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message ?? String.Empty : String.Empty);

                        JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Warn, exception);
                        Interlocked.Increment(ref this.failures);
                    }
                }
            }
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;

                JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Info, "EserviceNotification job disposed");
            }
        }

        private DateTime? CalculateNextSendingAttemptTime(int failedAttempts)
        {
            DateTime? nextAttemptTime = null;

            if (failedAttempts < 6)
            {
                nextAttemptTime = DateTime.Now.AddMinutes(1);
            }
            else if (failedAttempts < 10)
            {
                nextAttemptTime = DateTime.Now.AddMinutes(15);
            }
            else if (failedAttempts < 15)
            {
                nextAttemptTime = DateTime.Now.AddHours(1);
            }
            else if (failedAttempts < 21)
            {
                nextAttemptTime = DateTime.Now.AddHours(3);
            }
            else if (failedAttempts < 25)
            {
                nextAttemptTime = DateTime.Now.AddHours(6);
            }

            return nextAttemptTime;
        }

        private bool IsNotificationAccepted(HttpResponseMessage responseMessage, out string resultContent)
        {
            bool returnValue = false;
            resultContent = String.Empty;
            try
            {
                if (responseMessage.IsSuccessStatusCode)
                {
                    resultContent = responseMessage.Content.ReadAsStringAsync().Result;

                    JObject responseJson = JObject.Parse(resultContent);
                    returnValue = bool.Parse((string)responseJson.Properties().ToList().Single(e => e.Name.ToLower() == "success".ToLower()).Value);
                }

                return returnValue;
            }
            catch
            {
                if (resultContent.Length > 1000)
                {
                    resultContent = resultContent.Substring(0, 1000) + "... value is trimmed";
                }
                return returnValue;
            }
        }
    }
}
