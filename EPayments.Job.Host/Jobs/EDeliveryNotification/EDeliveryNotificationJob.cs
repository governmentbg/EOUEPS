using Autofac.Features.OwnedInstances;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
using EPayments.EDelivery.Manager;

namespace EPayments.Job.Host.Jobs.EDeliveryNotification
{
    public class EDeliveryNotificationJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, IDeliveryRegisterManager>>> dependencyFactory;
        private object syncRoot = new object();
        private bool disposed;
        private int batchSize;
        private int failedAttempts;
        private TimeSpan period;
        private int parallelTasks;
        private int successes;
        private int failures;

        private readonly string PaymentAisClientRequestCreated = AppSettings.EPaymentsJobHost_PaymentAisClientRequestCreated;
        private readonly string PaymentApplicantRequestCreated = AppSettings.EPaymentsJobHost_PaymentApplicantRequestCreated;
        private readonly string PaymentRequestAisClientShared = AppSettings.EPaymentsJobHost_PaymentRequestAisClientShared;
        private readonly string PaymentRequestApplicantClientShared = AppSettings.EPaymentsJobHost_PaymentRequestApplicantClientShared;
        private readonly string PaymentRequestAisClientObligationStatusChanged = AppSettings.EPaymentsJobHost_PaymentRequestAisClientObligationStatusChanged;
        private readonly string PaymentRequestApplicantObligationStatusChanged = AppSettings.EPaymentsJobHost_PaymentRequestApplicantObligationStatusChanged;
        private readonly string PaymentRequestAisClientObligationStatusCanceled = AppSettings.EPaymentsJobHost_PaymentRequestAisClientObligationStatusCanceled;
        private readonly string PaymentRequestApplicantObligationStatusCanceled = AppSettings.EPaymentsJobHost_PaymentRequestApplicantObligationStatusCanceled;
        private readonly string ServiceAddress = AppSettings.EPaymentsCommon_WebAddress;
        private readonly int Timeout = AppSettings.EPaymentsJobHost_EDeliveryNotificationTimeoutInMinutes;

        public EDeliveryNotificationJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, IDeliveryRegisterManager>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;

            this.failedAttempts = AppSettings.PaymentsJobHost_EDeliveryNotificationJobMaxFailedAttempts;
            this.batchSize = AppSettings.EPaymentsJobHost_EDeliveryNotificationJobBatchSize;
            this.period = TimeSpan.FromSeconds(AppSettings.EPaymentsJobHost_EserviceNotificationJobPeriodInSeconds);
            this.parallelTasks = AppSettings.EPaymentsJobHost_EserviceNotificationJobParallelTasks;
        }

        public string Name
        {
            get { return "EDeliveryNotificationJob"; }
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

                GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.EserviceNotificationJobLastInvocationTime);
                if (lastInvocationTime != null)
                {
                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();
                }
            }

            ICollection<int> pendingNotifications = new List<int>();

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
                    var deliveryManager = factory.Value.Item4;

                    pendingNotifications = jobRepository.GetPendingDeliveryEserviceNotificationIds(this.batchSize);
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
                        .Log(LogLevel.Info, string.Format("EDeliveryNotificationJob batch finished in {0}ms - {1} notifications send, {2} failures of total {3} notifications.", sw.ElapsedMilliseconds, this.successes, this.failures, pendingNotifications.Count));
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
                    .Log(LogLevel.Error, ex.Message, ex);
            }
        }

        private Task SendParallel(CancellationToken ct, ICollection<int> pendingNotificationIds)
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
            EserviceDeliveryNotification notification;
            PaymentRequestObligationLog notificationlog;
            PaymentRequest paymentRequest = null;

            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var jobRepository = factory.Value.Item2;
                var deliveryManager = factory.Value.Item4;

                while (notificationIds.TryDequeue(out notificationId))
                {
                    try
                    {
                        if (disposed)
                        {
                            break;
                        }

                        ct.ThrowIfCancellationRequested();

                        notificationlog = jobRepository.GetDeliveryEserviceNotificationById(notificationId);

                        // notifications are stopped for payment requests with algorithm 1 and 2 and obligation statuses { Paid, ForDistribution, CheckedAccount }
                        int algorithmId = notificationlog.PaymentRequest.ObligationType.AlgorithmId;
                        if ((algorithmId == 1 || algorithmId == 2) &&
                            (
                                notificationlog.ObligationStatusId == (int)ObligationStatusEnum.Paid ||
                                notificationlog.ObligationStatusId == (int)ObligationStatusEnum.ForDistribution ||
                                notificationlog.ObligationStatusId == (int)ObligationStatusEnum.CheckedAccount
                            )
                        )
                        {
                            JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Info, $"EDeliveryNotification has been detected for remove with PR-Id {notificationlog.PaymentRequestId} with algorithm {notificationlog.PaymentRequest.ObligationType.AlgorithmId} and status {notificationlog.ObligationStatusId}");
                            jobRepository.RemoveEDelivaryNotification(notificationlog);
                            continue;
                        }

                        // notifications are stopped for payment requests with algorithm 1 and 2 and request statuses different than Paid
                        if ((algorithmId == 1 || algorithmId == 2) && notificationlog.PaymentRequest.PaymentRequestStatusId != PaymentRequestStatus.Paid)
                        {
                            JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Info, $"EDeliveryNotification has been detected for remove with PR-Id {notificationlog.PaymentRequestId}");
                            jobRepository.RemoveEDelivaryNotification(notificationlog);
                            continue;
                        }

                        var notificationDefault = new EserviceDeliveryNotification(notificationlog, (int?)null);

                        notification = jobRepository.GetDeliveryEserviceNotificationByLogId(notificationId);

                        if (notification == null)
                        {
                            notification = notificationDefault;
                            jobRepository.AddEntity<EserviceDeliveryNotification>(notificationDefault);
                        }

                        paymentRequest = jobRepository.GetPaymentRequestById(notification.PaymentRequestId);

                        bool isInstitution = false;

                        if (paymentRequest != null)
                        {
                            if (paymentRequest.ApplicantUinTypeId == UinType.Bulstat)
                            {
                                isInstitution = true;
                            }
                        }

                        try
                        {
                            if (notification.NotificationStatusId == NotificationStatus.Pending || notification.NotificationStatusId == NotificationStatus.Error)
                            {
                                if (ServicePointManager.SecurityProtocol == (SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls))
                                {
                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                                }

                                try
                                {
                                    switch (notificationlog.ObligationStatusId)
                                    {
                                        case (int)ObligationStatusEnum.Asked:
                                            await this.SendMessages(notification,
                                                deliveryManager,
                                                string.Format(PaymentApplicantRequestCreated,
                                                notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                notificationlog.PaymentRequest.EserviceClient.AisName,
                                                CurrencyHelper.GetValueForDisplayUsingCurrencyMode(notificationlog.PaymentRequest.PaymentAmount, notificationlog.PaymentRequest.CreateDate),
                                                ServiceAddress,
                                                notificationlog.PaymentRequest.PaymentRequestAccessCode),
                                                notificationlog.PaymentRequest.EserviceClient.DeliveryAdminstrationId != null ?
                                                string.Format(PaymentAisClientRequestCreated,
                                                    notificationlog.PaymentRequest.PaymentReferenceNumber,
                                                    notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                    CurrencyHelper.GetValueForDisplayUsingCurrencyMode(notificationlog.PaymentRequest.PaymentAmount, notificationlog.PaymentRequest.CreateDate),
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.ExpirationDate)) : null,
                                                    isInstitution
                                                );
                                            break;
                                        case (int)ObligationStatusEnum.Ordered:
                                            await this.SendMessages(notification,
                                                deliveryManager,
                                                string.Format(PaymentRequestApplicantClientShared,
                                                notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                ObligationStatusEnum.Ordered.GetDescription(),
                                                CurrencyHelper.GetValueForDisplayUsingCurrencyMode(notificationlog.PaymentRequest.PaymentAmount, notificationlog.PaymentRequest.CreateDate),
                                                    notificationlog.PaymentRequest.PaymentRequestAccessCode,
                                                ServiceAddress),
                                                notificationlog.PaymentRequest.EserviceClient.DeliveryAdminstrationId != null ?
                                                String.Format(PaymentRequestAisClientShared,
                                                    notificationlog.PaymentRequest.PaymentReferenceNumber,
                                                    notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                    ObligationStatusEnum.Ordered.GetDescription(),
                                                    CurrencyHelper.GetValueForDisplayUsingCurrencyMode(notificationlog.PaymentRequest.PaymentAmount, notificationlog.PaymentRequest.CreateDate),
                                                    notificationlog.PaymentRequest.PaymentRequestAccessCode) :
                                                    null,
                                                    isInstitution
                                                );
                                            break;
                                        case (int)ObligationStatusEnum.Canceled:
                                            await this.SendMessages(notification,
                                                    deliveryManager,
                                                    string.Format(PaymentRequestApplicantObligationStatusCanceled,
                                                    notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                    ObligationStatusEnum.Canceled.GetDescription(),
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                    ServiceAddress),
                                                    notificationlog.PaymentRequest.EserviceClient.DeliveryAdminstrationId != null ?
                                                    string.Format(PaymentRequestAisClientObligationStatusCanceled,
                                                    notificationlog.PaymentRequest.PaymentReferenceNumber,
                                                    notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                    ObligationStatusEnum.Canceled.GetDescription()) :
                                                    null,
                                                    isInstitution
                                                );
                                            break;
                                        case (int)ObligationStatusEnum.IrrevocableOrder:
                                            await this.SendMessages(notification,
                                                    deliveryManager,
                                                    string.Format(PaymentRequestApplicantObligationStatusChanged,
                                                    notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                    Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                    ObligationStatusEnum.IrrevocableOrder.GetDescription(),
                                                    CurrencyHelper.GetValueForDisplayUsingCurrencyMode(notificationlog.PaymentRequest.PaymentAmount, notificationlog.PaymentRequest.CreateDate),
                                                    ServiceAddress),
                                                    notificationlog.PaymentRequest.EserviceClient.DeliveryAdminstrationId != null ?
                                                    string.Format(PaymentRequestAisClientObligationStatusChanged,
                                                        notificationlog.PaymentRequest.PaymentReferenceNumber,
                                                        notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                        Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                        Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                        ObligationStatusEnum.IrrevocableOrder.GetDescription()) :
                                                        null,
                                                    isInstitution
                                                );
                                            break;
                                        case (int)ObligationStatusEnum.Paid:
                                        case (int)ObligationStatusEnum.CheckedAccount:
                                        case (int)ObligationStatusEnum.ForDistribution:
                                            if (notificationlog.PaymentRequest.EserviceClient.DeliveryAdminstrationId != null)
                                            {
                                                string aisClientMessage = string.Format(PaymentRequestAisClientObligationStatusChanged,
                                                notificationlog.PaymentRequest.PaymentReferenceNumber,
                                                notificationlog.PaymentRequest.PaymentRequestIdentifier,
                                                Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.PaymentRequest.CreateDate),
                                                Formatter.DateTimeToBgFormatNotLocalTime(notificationlog.ChangeDate),
                                                ((ObligationStatusEnum)notificationlog.ObligationStatusId).GetDescription());

                                                int taskAisClient = await deliveryManager.SendMessageAsync(notification, aisClientMessage);

                                                notification.AddResponseCodes(taskAisClient.ToString());
                                            }
                                            else
                                            {
                                                notification.NotificationStatusId = NotificationStatus.Terminated;
                                            }

                                            break;
                                        default:
                                            throw new Exception("Invalid message option");
                                    }

                                    notification.SetStatus(NotificationStatus.Sent);
                                    Interlocked.Increment(ref this.successes);
                                }
                                catch (AggregateException aggrEx)
                                {
                                    string exceptionMessage = aggrEx.Message;

                                    var exception = "DeliveryService: " + (exceptionMessage ?? String.Empty) + " " +
                                        string.Join(", ", aggrEx.InnerExceptions
                                            .Where(ex => ex != null)
                                            .Select(ex => String.Format(" Inner exception: {0}, StackTrace ->{1}", ex.Message, ex.StackTrace)));

                                    notification.SetStatus(NotificationStatus.Error);
                                    notification.IncrementFailedAttempts(exception);
                                    DateTime? sendNotBefore = CalculateNextSendingAttemptTime(notification.FailedAttempts);
                                    notification.SetNextSendingAttemptTime(sendNotBefore, !sendNotBefore.HasValue);

                                    JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Warn, exception, aggrEx);
                                    Interlocked.Increment(ref this.failures);
                                }
                                catch (Exception ex)
                                {
                                    var exception = "DeliveryService: " + (ex.Message ?? String.Empty) + (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message ?? String.Empty : String.Empty);

                                    notification.SetStatus(NotificationStatus.Error);
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
                            notification.IncrementFailedAttempts(ex.Message);

                            JobLogger.Get(JobName.EserviceNotification).Log(LogLevel.Error, ex.Message, ex);
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

        private async Task SendMessages(EserviceDeliveryNotification notification,
            IDeliveryRegisterManager deliveryManager,
            string applicantMessage,
            string aisClientMessage,
            bool isInstitution)
        {
            Exception aisClientException = null;
            string client = "client";
            string user = "user";

            if ((notification.Uniqueidentifier != null && aisClientMessage != null && notification.FailedAttempts == 0) ||
                (notification.Uniqueidentifier != null && aisClientMessage != null && notification.FailedAttempts > 0 && (string.IsNullOrWhiteSpace(notification.ResponseCodes) || !notification.ResponseCodes.Contains(client))))
            {
                try
                {
                    int taskAisClient = await deliveryManager.SendMessageAsync(notification, aisClientMessage);

                    notification.ResponseCodes += CreateResponseMessage(client, taskAisClient.ToString());
                }
                catch (Exception ex)
                {
                    if (string.IsNullOrWhiteSpace(notification.PersonUniqueIdentifier) ||
                        (notification.FailedAttempts > 0 && (!string.IsNullOrWhiteSpace(notification.ResponseCodes) && notification.ResponseCodes.Contains(user))))
                    {
                        throw;
                    }

                    aisClientException = ex;
                }
            }

            if (!string.IsNullOrWhiteSpace(notification.PersonUniqueIdentifier) &&
                (string.IsNullOrWhiteSpace(notification.ResponseCodes) || !notification.ResponseCodes.Contains(user)))
            {
                try
                {
                    int taskPerson = await deliveryManager.SendMessagePersonAsync(notification, applicantMessage, isInstitution);

                    notification.ResponseCodes += CreateResponseMessage(user, taskPerson.ToString());
                }
                catch (Exception ex)
                {
                    if (aisClientException == null)
                    {
                        throw;
                    }

                    throw new AggregateException(aisClientException, ex);
                }
            }

            if (aisClientException != null)
            {
                throw aisClientException;
            }
        }

        private string CreateResponseMessage(string prefix, string postfix)
        {
            return string.Format("{0}={1};", prefix, postfix);
        }

        private DateTime? CalculateNextSendingAttemptTime(int failedAttempts)
        {
            DateTime? nextAttemptTime = null;

            if (failedAttempts < this.failedAttempts)
            {
                nextAttemptTime = DateTime.Now.AddMinutes(Timeout);
            }

            return nextAttemptTime;
        }
    }
}
