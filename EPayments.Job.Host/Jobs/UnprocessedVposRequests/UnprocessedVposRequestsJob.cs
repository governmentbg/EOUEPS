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
using EPayments.Common.VposHelpers;
using Newtonsoft.Json;

namespace EPayments.Job.Host.ExpiredRequest
{
    public class UnprocessedVposRequestsJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory;
        private bool disposed;
        private TimeSpan period;
        
        private readonly int fiBankCheckPendingRequests_MaxFailedAttempts =
            AppSettings.EPaymentsJobHost_UnprocessedVposRequestsJobFiBankMaxFailedAttempts;

        private readonly TimeSpan fiBankCheckPendingRequests_FailedAttemptTimeout =
            TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_UnprocessedVposRequestsJobFiBankFailedAttemptTimeoutInMinutes);

        private readonly int fiBankCheckPendingRequests_FibankRequestTimeoutInMinutes =
            AppSettings.EPaymentsJobHost_UnprocessedVposRequestsJobFibankRequestTimeoutInMinutes;

        private readonly string fiBankVposMerchantHandlerUrl =
            AppSettings.EPaymentsJobHost_UnprocessedVposRequestsJobFiBankVposMerchantHandlerUrl;

        public UnprocessedVposRequestsJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;

            this.period = TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_UnprocessedVposRequestsJobPeriodInMinutes);
        }

        public string Name
        {
            get { return "UnprocessedVposRequestsJob"; }
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

            try
            {
                if (disposed)
                {
                    return;
                }

                ProcessPendingVposFiBankRequest(ct);
                ProcessPendingDskEcommRequest(ct);
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.ExpiredRequest)
                    .Log(LogLevel.Error, "Job was canceled due to a token cancellation request", ex);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.ExpiredRequest)
                    .Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
            }
        }

        public void ProcessPendingVposFiBankRequest(CancellationToken ct)
        {
            while (true)
            {
                VposFiBankRequest pendingRequest = null;

                using (var factory = this.dependencyFactory())
                {
                    var unitOfWork = factory.Value.Item1;
                    var jobRepository = factory.Value.Item2;
                    var systemRepository = factory.Value.Item3;

                    pendingRequest = jobRepository.GetPendingResultVposFiBankRequest(
                        fiBankCheckPendingRequests_MaxFailedAttempts,
                        fiBankCheckPendingRequests_FailedAttemptTimeout,
                        fiBankCheckPendingRequests_FibankRequestTimeoutInMinutes);

                    if (pendingRequest == null)
                    {
                        break;
                    }
                    else
                    {
                        Tuple<FiBankVposTransactionResult, string> transactionResult = null;
                        Exception exception = null;

                        try
                        {
                            transactionResult = FiBankVposService.CheckTransactionResult(
                                fiBankVposMerchantHandlerUrl,
                                pendingRequest.ClientIpAddress,
                                pendingRequest.PaymentRequest.EserviceClient.FiBankVposAccessKeystoreFilename,
                                pendingRequest.PaymentRequest.EserviceClient.FiBankVposAccessKeystorePassword,
                                pendingRequest.TransactionId);
                        }
                        catch (Exception ex)
                        {
                            exception = ex;
                        }

                        pendingRequest.JobLastCheckResultDate = DateTime.Now;
                        pendingRequest.JobLastCheckResultTransactionInfo = transactionResult != null ? transactionResult.Item2 : null;

                        if (transactionResult != null && transactionResult.Item1 != FiBankVposTransactionResult.Pending)
                        {
                            pendingRequest.ResultStatus = VposRequestResultStatus.ReceivedByCheckStatusJob;
                            pendingRequest.IsPaymentSuccessful = transactionResult.Item1 == FiBankVposTransactionResult.Successful;

                            if (transactionResult.Item1 == FiBankVposTransactionResult.Successful)
                            {
                                MarkPaymentRequestAsPaid(pendingRequest.PaymentRequest, pendingRequest.TransactionId, systemRepository);
                            }
                        }
                        else
                        {
                            var errorData = new
                            {
                                TransactionResultReceived = transactionResult != null,
                                FiBankVposTransactionResult = transactionResult != null ? transactionResult.Item1.ToString() : null,
                                ResultData = transactionResult != null ? transactionResult.Item2 : null,
                                Exception = exception != null ? Formatter.ExceptionToDetailedInfo(exception) : null
                            };

                            pendingRequest.JobCheckResultFailedAttempts = pendingRequest.JobCheckResultFailedAttempts + 1;
                            pendingRequest.JobCheckResultFailedAttemptsErrors =
                                (pendingRequest.JobCheckResultFailedAttemptsErrors != null ? pendingRequest.JobCheckResultFailedAttemptsErrors + Environment.NewLine : String.Empty) +
                                JsonConvert.SerializeObject(errorData);
                        }

                        unitOfWork.Save();
                    }
                }

                if (ct.IsCancellationRequested)
                    break;
            }
        }

        public void ProcessPendingDskEcommRequest(CancellationToken ct)
        {
            while (true)
            {
                VposDskEcommRequest pendingRequest = null;

                using (var factory = this.dependencyFactory())
                {
                    var unitOfWork = factory.Value.Item1;
                    var jobRepository = factory.Value.Item2;
                    var systemRepository = factory.Value.Item3;

                    pendingRequest = jobRepository.GetPendingResultVposDskEcommRequest(
                        fiBankCheckPendingRequests_MaxFailedAttempts,
                        fiBankCheckPendingRequests_FailedAttemptTimeout,
                        fiBankCheckPendingRequests_FibankRequestTimeoutInMinutes);

                    if (pendingRequest == null)
                    {
                        break;
                    }
                    else
                    {
                        Tuple<DskEcommVposOrderStatus, string> transactionResult = null;
                        Exception exception = null;

                        try
                        {
                            transactionResult = DskEcommVposService.CheckOrderStatus(
                                pendingRequest.PaymentRequest.EserviceClient.VposClient.PaymentRequestUrl,
                                pendingRequest.PaymentRequest.EserviceClient.DskVposMerchantId,
                                pendingRequest.PaymentRequest.EserviceClient.DskVposMerchantPassword,
                                pendingRequest.TransactionId);
                        }
                        catch (Exception ex)
                        {
                            exception = ex;
                        }

                        pendingRequest.JobLastCheckResultDate = DateTime.Now;
                        pendingRequest.JobLastCheckResultTransactionInfo = transactionResult != null ? transactionResult.Item2 : null;

                        if (transactionResult != null && transactionResult.Item1.OrderStatus.Value != DskEcommVposTransactionResult.Pending)
                        {
                            pendingRequest.ResultStatus = VposRequestResultStatus.ReceivedByCheckStatusJob;
                            pendingRequest.IsPaymentSuccessful = transactionResult.Item1.OrderStatus.Value == DskEcommVposTransactionResult.Successful;

                            if (transactionResult.Item1.OrderStatus.Value == DskEcommVposTransactionResult.Successful)
                            {
                                MarkPaymentRequestAsPaid(pendingRequest.PaymentRequest, pendingRequest.TransactionId, systemRepository);
                            }
                        }
                        else
                        {
                            var errorData = new
                            {
                                TransactionResultReceived = transactionResult != null,
                                FiBankVposTransactionResult = transactionResult != null ? transactionResult.Item1.ToString() : null,
                                ResultData = transactionResult != null ? transactionResult.Item2 : null,
                                Exception = exception != null ? Formatter.ExceptionToDetailedInfo(exception) : null
                            };

                            pendingRequest.JobCheckResultFailedAttempts = pendingRequest.JobCheckResultFailedAttempts + 1;
                            pendingRequest.JobCheckResultFailedAttemptsErrors =
                                (pendingRequest.JobCheckResultFailedAttemptsErrors != null ? pendingRequest.JobCheckResultFailedAttemptsErrors + Environment.NewLine : String.Empty) +
                                JsonConvert.SerializeObject(errorData);
                        }

                        unitOfWork.Save();
                    }
                }

                if (ct.IsCancellationRequested)
                    break;
            }
        }

        private void MarkPaymentRequestAsPaid(PaymentRequest paymentRequest, string transactionId, ISystemRepository systemRepository)
        {
            paymentRequest.IsVposAuthorized = true;
            paymentRequest.VposAuthorizationId = transactionId;
            paymentRequest.PaymentRequestStatusId = Model.Enums.PaymentRequestStatus.Authorized;
            paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

            User user = systemRepository.GetUserByUin(paymentRequest.ApplicantUin);

            if (user != null && !String.IsNullOrWhiteSpace(user.Email))
            {
                if (user.StatusNotifications)
                {
                    StatusChangedPaymentRequestContextDO contextDO = new StatusChangedPaymentRequestContextDO(
                        paymentRequest.PaymentRequestIdentifier,
                        paymentRequest.ServiceProviderName,
                        paymentRequest.PaymentReason,
                        paymentRequest.PaymentAmount,
                        paymentRequest.PaymentRequestStatusId.GetDescription(),
                        paymentRequest.CreateDate);

                    EPayments.Model.Models.Email email = new EPayments.Model.Models.Email(contextDO, user.Email);

                    systemRepository.AddEntity<EPayments.Model.Models.Email>(email);
                }

                if (user.StatusObligationNotifications)
                {
                    StatusChangedObligationContextDO contextOblDO = new StatusChangedObligationContextDO(
                        paymentRequest.PaymentRequestIdentifier,
                        paymentRequest.ServiceProviderName,
                        paymentRequest.PaymentReason,
                        paymentRequest.PaymentAmount,
                        paymentRequest.ObligationStatusId.GetDescription(),
                        paymentRequest.CreateDate);

                    EPayments.Model.Models.Email email = new EPayments.Model.Models.Email(contextOblDO, user.Email);

                    systemRepository.AddEntity<EPayments.Model.Models.Email>(email);
                }
            }
            
            if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
            {
                EPayments.Model.Models.EserviceNotification statusNotification = new EPayments.Model.Models.EserviceNotification(paymentRequest, (int?)null);

                systemRepository.AddEntity<EPayments.Model.Models.EserviceNotification>(statusNotification);
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
