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
using System.Runtime.Remoting.Contexts;

namespace EPayments.Job.Host.ExpiredRequest
{
    public class ExpiredRequestJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory;
        private bool disposed;
        private TimeSpan period;
        private int successes;
        private int failures;

        public ExpiredRequestJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;

            this.period = TimeSpan.FromSeconds(AppSettings.EPaymentsJobHost_ExpiredRequestJobPeriodInSeconds);
        }

        public string Name
        {
            get { return "ExpiredRequestJob"; }
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

                GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.ExpiredRequestJobLastInvocationTime);
                if (lastInvocationTime != null)
                {
                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();
                }
            }

            IList<int> expiredRequestIds = new List<int>();

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

                    expiredRequestIds = jobRepository.GetExpiredRequestIds();
                }

                if (expiredRequestIds.Any())
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();

                    this.successes = 0;
                    this.failures = 0;

                    ConcurrentQueue<int> requestIds = new ConcurrentQueue<int>(expiredRequestIds);

                    SetAsExpired(ct, requestIds);

                    sw.Stop();

                    JobLogger.Get(JobName.ExpiredRequest)
                        .Log(LogLevel.Info, string.Format("ExpiredRequest batch finished in {0}ms - {1} successes, {2} failures of total {3} notifications.", sw.ElapsedMilliseconds, this.successes, this.failures, expiredRequestIds.Count));
                }

                AddEmailsForCertificateExpirationNotifications(ct);
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.ExpiredRequest)
                    .Log(LogLevel.Error, string.Format("Job was canceled due to a token cancellation request; ExpiredRequest batch finished with {0} successes, {1} failures of total {2} notifications.", this.successes, this.failures, expiredRequestIds.Count), ex);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.ExpiredRequest)
                    .Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
            }
        }

        private void SetAsExpired(CancellationToken ct, ConcurrentQueue<int> requestIds)
        {
            int requestId;

            while (requestIds.TryDequeue(out requestId))
            {
                if (disposed)
                {
                    break;
                }

                try
                {
                    using (var factory = this.dependencyFactory())
                    {
                        var unitOfWork = factory.Value.Item1;
                        var jobRepository = factory.Value.Item2;
                        var systemRepository = factory.Value.Item3;

                        ct.ThrowIfCancellationRequested();

                        PaymentRequest paymentRequest = jobRepository.GetRequestById(requestId);

                        if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Paid 
                            || paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Canceled
                            || paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Suspended)
                        {
                            return;
                        }
                        
                        paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Expired;
                        paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                        //Send email if notification is turned on
                        EPayments.Model.Models.User user = systemRepository.GetUserByUin(paymentRequest.ApplicantUin);
                        if (user != null && user.StatusNotifications && !String.IsNullOrWhiteSpace(user.Email))
                        {
                            StatusChangedPaymentRequestContextDO contextDO = new StatusChangedPaymentRequestContextDO(
                                paymentRequest.PaymentRequestIdentifier,
                                paymentRequest.ServiceProviderName,
                                paymentRequest.PaymentReason,
                                paymentRequest.PaymentAmount,
                                paymentRequest.PaymentRequestStatusId.GetDescription(),
                                paymentRequest.CreateDate);

                            EPayments.Model.Models.Email email = new EPayments.Model.Models.Email(contextDO, user.Email);

                            systemRepository.AddEntity(email);
                        }

                        //Send status change notification if notification is turned on
                        if (!(paymentRequest.ObligationType != null && paymentRequest.ObligationType.AlgorithmId == 2) 
                            && !String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                        {
                            Model.Models.EserviceNotification statusNotification = new EPayments.Model.Models.EserviceNotification(paymentRequest);

                            systemRepository.AddEntity(statusNotification);
                        }

                        unitOfWork.Save();

                        this.successes++;
                    }
                }
                catch (Exception ex)
                {
                    this.failures++;
                    JobLogger.Get(JobName.ExpiredRequest).Log(LogLevel.Error, String.Format("Error occured on paymentRequestId {0}. Message: {1}, StackTrace: {2}", requestId, ex.Message, ex.StackTrace), ex);
                }
            }
        }

        private void AddEmailsForCertificateExpirationNotifications(CancellationToken ct)
        {
            if (disposed)
            {
                return;
            }

            int? eserviceClientId = null;

            try
            {
                using (var factory = this.dependencyFactory())
                {
                    var unitOfWork = factory.Value.Item1;
                    var jobRepository = factory.Value.Item2;
                    var systemRepository = factory.Value.Item3;

                    ct.ThrowIfCancellationRequested();

                    IList<EserviceClient> clients = jobRepository.GetAllActiveEserviceClients();

                    foreach(var client in clients)
                    {
                        eserviceClientId = client.EserviceClientId;

                        if (client.BoricaVposRequestSignCertValidTo.HasValue &&
                            DateTime.Now.AddDays(20) >= client.BoricaVposRequestSignCertValidTo.Value && 
                            !(client.BoricaVposRequestSignCertExpMailSend.HasValue && client.BoricaVposRequestSignCertExpMailSend.Value))
                        {
                            CertificateExpirationContextDO contextDO = new CertificateExpirationContextDO(
                            client.EserviceClientId,
                            client.AisName,
                            client.BoricaVposRequestSignCertFileName,
                            client.BoricaVposRequestSignCertValidTo.Value,
                            true);

                            EPayments.Model.Models.Email email = new EPayments.Model.Models.Email(contextDO, AppSettings.EPaymentsJobHost_ExpiredRequestJobFeedbackEmail);

                            jobRepository.AddEntity<EPayments.Model.Models.Email>(email);

                            client.BoricaVposRequestSignCertExpMailSend = true;
                        }
                    }

                    DateTime boricaVposResponseSignCertValidTo = Parser.BgFormatDateStringToDateTime(systemRepository.GetGlobalValueByKey(GlobalValueKey.BoricaVposResponseSignCertValidTo).Value).Value;
                    GlobalValue boricaVposResponseSignCertExpMailSendGlobalValue = systemRepository.GetGlobalValueByKey(GlobalValueKey.BoricaVposResponseSignCertExpMailSend);

                    if (DateTime.Now.AddDays(20) >= boricaVposResponseSignCertValidTo && boricaVposResponseSignCertExpMailSendGlobalValue.Value != "1")
                    {
                        boricaVposResponseSignCertExpMailSendGlobalValue.Value = "1";

                        CertificateExpirationContextDO contextDO = new CertificateExpirationContextDO(
                            -1,
                            "GLOBAL BORICA RESPONSE CERTIFICATE",
                            "GLOBAL BORICA RESPONSE CERTIFICATE",
                            boricaVposResponseSignCertValidTo,
                            false);

                        EPayments.Model.Models.Email email = new EPayments.Model.Models.Email(contextDO, AppSettings.EPaymentsJobHost_ExpiredRequestJobFeedbackEmail);
                        jobRepository.AddEntity<EPayments.Model.Models.Email>(email);
                    }

                    unitOfWork.Save();
                }
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.ExpiredRequest).Log(LogLevel.Error, String.Format("Error occured on AddEmailsForCertificateExpirationNotifications for EserviceClient with id {0}. Message: {1}", eserviceClientId.HasValue ? eserviceClientId.ToString() : String.Empty, ex.Message), ex);
            }
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;

                JobLogger.Get(JobName.ExpiredRequest).Log(LogLevel.Info, "ExpiredRequest job disposed");
            }
        }
    }
}
