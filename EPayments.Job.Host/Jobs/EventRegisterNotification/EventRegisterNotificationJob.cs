using Autofac.Features.OwnedInstances;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using EPayments.Job.Host.Core;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using NLog;
using EPayments.Model.DataObjects.EmailTemplateContext;
using EPayments.Common.Helpers;
using EPayments.Model.Models;
using EPayments.EventRegister.Manager;
using EPayments.EventRegister.DataObjects;
using EPayments.EventRegister.EventRegisterService;

namespace EPayments.Job.Host.EventRegisterNotification
{
    public class EventRegisterNotificationJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, IEventRegisterManager>>> dependencyFactory;
        private bool disposed;
        private TimeSpan period;
        private int successes;
        private int failures;

        public EventRegisterNotificationJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, IEventRegisterManager>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;

            this.period = TimeSpan.FromSeconds(AppSettings.EPaymentsJobHost_EventRegisterNotificationJobPeriodInSeconds);
        }

        public string Name
        {
            get { return "EventRegisterNotificationJob"; }
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

                GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.EventRegisterNotificationJobLastInvocationTime);
                if (lastInvocationTime != null)
                {
                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();
                }
            }

            IList<int> eventRegisterIds = new List<int>();

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

                    eventRegisterIds = jobRepository.GetPendingEventRegisterNotificationIds();
                }

                if (eventRegisterIds.Any())
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();

                    this.successes = 0;
                    this.failures = 0;

                    ConcurrentQueue<int> ids = new ConcurrentQueue<int>(eventRegisterIds);

                    SendDataToEventRegister(ct, ids);

                    sw.Stop();

                    JobLogger.Get(JobName.EventRegisterNotification)
                        .Log(LogLevel.Info, string.Format("EventRegister batch finished in {0}ms - {1} successes, {2} failures of total {3} notifications.", sw.ElapsedMilliseconds, this.successes, this.failures, eventRegisterIds.Count));
                }
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.EventRegisterNotification)
                    .Log(LogLevel.Error, string.Format("Job was canceled due to a token cancellation request; EventRegister batch finished with {0} successes, {1} failures of total {2} notifications.", this.successes, this.failures, eventRegisterIds.Count), ex);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.EventRegisterNotification)
                    .Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
            }
        }

        private void SendDataToEventRegister(CancellationToken ct, ConcurrentQueue<int> ids)
        {
            int erId;

            while (ids.TryDequeue(out erId))
            {
                if (disposed)
                {
                    break;
                }

                using (var factory = this.dependencyFactory())
                {
                    var unitOfWork = factory.Value.Item1;
                    var jobRepository = factory.Value.Item2;
                    var systemRepository = factory.Value.Item3;
                    var eventRegisterManager = factory.Value.Item4;

                    ct.ThrowIfCancellationRequested();

                    EPayments.Model.Models.EventRegisterNotification notification = jobRepository.GetEventRegisterNotificationById(erId);

                    try
                    {
                        EventDO eventDO = new EventDO();
                        eventDO.DocumentRegNumber = notification.EventDocRegNumber;
                        eventDO.EventDescription = notification.EventDescription;
                        eventDO.Time = notification.EventTime;
                        switch ((EventRegisterNotificationType)Enum.Parse(typeof(EventRegisterNotificationType), notification.EventType))
                        {
                            case EventRegisterNotificationType.PaymentRequestRegistered:
                                eventDO.EventType = EventTypeEnum.EPAYMENT_PAYMENT_REQUEST_RECEIVED;
                                break;
                            case EventRegisterNotificationType.PaymentRequestDenied:
                                eventDO.EventType = EventTypeEnum.EPAYMENT_PAYMENT_REQUEST_DENIED;
                                break;
                            case EventRegisterNotificationType.VposPaymentAuthorized:
                                eventDO.EventType = EventTypeEnum.EPAYMENT_PAYMENT_VPOS_AUTHORIZED;
                                break;
                            default:
                                throw new Exception("Invalid EventRegisterNotification EventType.");
                        }

                        try
                        {
                            eventRegisterManager.LogEventAsync(eventDO).Wait();

                            notification.SetStatus(NotificationStatus.Sent);
                            this.successes++;
                        }
                        catch (Exception ex)
                        {
                            var exception = "EventRegisterManager LogEventAsync exception: " + (ex.Message ?? String.Empty) + (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message ?? String.Empty : String.Empty);

                            notification.IncrementFailedAttempts(exception);
                            DateTime? sendNotBefore = CalculateNextSendingAttemptTime(notification.FailedAttempts);
                            notification.SetNextSendingAttemptTime(sendNotBefore, !sendNotBefore.HasValue);

                            JobLogger.Get(JobName.EventRegisterNotification).Log(LogLevel.Warn, exception);
                            this.failures++;
                        }
                    }
                    catch (Exception ex)
                    {
                        notification.SetStatus(NotificationStatus.Error);
                        notification.IncrementFailedAttempts($"{ex.Message}, StackTrace -> {ex.StackTrace}");

                        JobLogger.Get(JobName.EventRegisterNotification).Log(LogLevel.Error, String.Format("Error occured on eventRegisterId {0}. Message: {1}, StackTrace: {2}", erId, ex.Message, ex.StackTrace), ex);
                        this.failures++;
                    }

                    unitOfWork.Save();
                }
            }
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;

                JobLogger.Get(JobName.EventRegisterNotification).Log(LogLevel.Info, "EventRegisterNotification job disposed");
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
    }
}
