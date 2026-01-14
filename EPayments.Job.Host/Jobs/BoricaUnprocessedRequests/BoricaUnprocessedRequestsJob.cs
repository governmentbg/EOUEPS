using Autofac.Features.OwnedInstances;
using EPayments.Common;
using EPayments.Common.BoricaHelpers;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Admin;
using EPayments.Data.ViewObjects.Web.APGModels;
using EPayments.Data.ViewObjects.Web.APGModels.Requests;
using EPayments.Job.Host.Core;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using Mail = EPayments.Model.Models.Email;
using Notification = EPayments.Model.Models.EserviceNotification;

namespace EPayments.Job.Host.Jobs.BoricaUnprocessedRequests
{
    public class BoricaUnprocessedRequestsJob : IJob
    {
        private bool _isDisposed = false;

        private Func<Owned<DisposableTuple<IUnitOfWork, ISystemRepository>>> DependencyFactory;

        public BoricaUnprocessedRequestsJob(Func<Owned<DisposableTuple<IUnitOfWork, ISystemRepository>>> dependencyFactory)
        {
            DependencyFactory = dependencyFactory;
        }

        public string Name => nameof(BoricaUnprocessedRequestsJob);

        public TimeSpan Period => TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_BoricaUnprocessedRequestsJobPeriodInMinutes);

        public void Action(CancellationToken token)
        {
            DisposableTuple<IUnitOfWork, ISystemRepository> value = DependencyFactory().Value;

            IUnitOfWork unitOfWork = value.Item1;
            ISystemRepository systemRepository = value.Item2;

            GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.BoricaUnprocessedRequestsJobLastInvocationTime);

            if (lastInvocationTime != null)
            {
                lastInvocationTime.ModifyDate = DateTime.Now;
            }
            else
            {
                systemRepository.AddEntity(new GlobalValue()
                {
                    Key = GlobalValueKey.BoricaUnprocessedRequestsJobLastInvocationTime.ToString(),
                    Value = null,
                    ModifyDate = DateTime.Now
                });
            }

            unitOfWork.Save();

            int limitInMinutes = AppSettings.EPaymentsJobHost_BoricaUnprocessedRequestsLimitTimeSpanInMinutes;
            int timeToWait = AppSettings.EPaymentsJobHost_BoricaUnprocessedRequestsWaitAfterEachRequestSeconds * 1000;
            int transactionsToTake = AppSettings.EPaymentsJobHost_BoricaUnprocessedRequestsTransactionsToTake;
            DateTime transactionStartDate = DateTime.Now;
            int cancelPayment = (int)BoricaTransactionStatusEnum.Canceled;
            int pendingPayment = (int)BoricaTransactionStatusEnum.Pending;
            int finalLimit = AppSettings.EPaymentsJobHost_BoricaFinalLimitInHours;

            try
            {
                List<BoricaTransaction> boricaTransactions = unitOfWork.DbContext.Set<BoricaTransaction>()
                    .Include(bt => bt.PaymentRequests.Select(pr => pr.ObligationType))
                    .Where(bt => bt.TransactionStatusId == pendingPayment)
                    .OrderByDescending(bt => bt.BoricaTransactionId)
                    .Take(transactionsToTake)
                    .ToList();

                if (boricaTransactions != null && boricaTransactions.Count > 0)
                {
                    UpdateNetProtocol();

                    foreach (BoricaTransaction boricaTransaction in boricaTransactions)
                    {
                        Thread.Sleep(timeToWait);

                        try
                        {
                            APGWStatusCheckRequestDataVO apgwRequest = CreateBoricaRequestData(boricaTransaction);

                            string content = null;

                            HttpResponseMessage response = BoricaRetryPolicy.GetBoricaRetryPolicy(JobName.BoricaUnprocessedRequestsJob).Execute(() =>
                            {
                                using (HttpClient httpClient = new HttpClient())
                                {
                                    return httpClient.PostAsync(apgwRequest.PostUrl, new FormUrlEncodedContent(apgwRequest.RequestFields())).Result;
                                }
                            });

                            if (response.StatusCode == HttpStatusCode.OK)
                            {
                                content = response.Content.ReadAsStringAsync().Result;

                                Dictionary<string, object> values = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                                APGWPaymentResponseDataDO apgwResponse = new APGWPaymentResponseDataDO();
                                apgwResponse.ParseFromDictionary(values);
                                JobLogger.Get(JobName.BoricaUnprocessedRequestsJob).Log(LogLevel.Info, $"borica transaction id{boricaTransaction.BoricaTransactionId} response: {JsonConvert.SerializeObject(apgwResponse)}");

                                if (VerifyResponse(apgwResponse))
                                {
                                    if (apgwResponse.Action == 0 && apgwResponse.Rc == "00")
                                    {
                                        apgwResponse.UpdateBoricaTransactionSuccessPayment(boricaTransaction);

                                        string[] applicantUins = boricaTransaction.PaymentRequests
                                            .Where(pr => !string.IsNullOrWhiteSpace(pr.ApplicantUin))
                                            .Select(pr => pr.ApplicantUin)
                                            .ToArray();

                                        User[] users = unitOfWork.DbContext.Set<User>()
                                            .Where(u => applicantUins.Contains(u.Egn))
                                            .ToArray();

                                        foreach (var paymentRequest in boricaTransaction.PaymentRequests)
                                        {
                                            if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Paid)
                                            {
                                                continue;
                                            }

                                            paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Paid;
                                            paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                                            if (paymentRequest.ObligationStatusId == null
                                                || paymentRequest.ObligationStatusId == ObligationStatusEnum.Ordered
                                                || paymentRequest.ObligationStatusId == ObligationStatusEnum.Asked)
                                            {
                                                paymentRequest.ObligationStatusId = ObligationStatusEnum.IrrevocableOrder;
                                            }

                                            User user = users.FirstOrDefault(u => string.Equals(u.Egn, paymentRequest.ApplicantUin, StringComparison.OrdinalIgnoreCase));

                                            if (user != null && !String.IsNullOrWhiteSpace(user.Email))
                                            {
                                                if (user.StatusObligationNotifications)
                                                {
                                                    unitOfWork.DbContext.Set<Mail>().Add(user.CreateStatusObligationNotificationEmail(paymentRequest));
                                                }

                                                if (user.StatusNotifications)
                                                {
                                                    unitOfWork.DbContext.Set<Mail>().Add(user.CreateStatusNotificationEmail(paymentRequest));
                                                }
                                            }

                                            if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                                            {
                                                Notification statusNotification = new Notification(paymentRequest, (int?)null);

                                                unitOfWork.DbContext.Set<Notification>().Add(statusNotification);
                                            }
                                        }

                                        unitOfWork.Save();
                                    }
                                    else if (apgwResponse.Rc != "00" && !apgwResponse.Rc.StartsWith("-"))
                                    {
                                        foreach (var paymentRequest in boricaTransaction.PaymentRequests)
                                        {
                                            if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess)
                                            {
                                                paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Pending;
                                                paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                                                if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                                                {
                                                    Notification statusNotification = new Notification(paymentRequest, (int?)null);

                                                    unitOfWork.DbContext.Set<Notification>().Add(statusNotification);
                                                }
                                            }
                                        }
                                        apgwResponse.UpdateBoricaTransactionCanceledPayment(boricaTransaction);
                                        unitOfWork.Save();
                                    }
                                    else if (apgwRequest.TRAN_TRTYPE == TrTypeEnum.Check && boricaTransaction.TransactionDate.AddMinutes(20) <= DateTime.Now)
                                    {
                                        foreach (var paymentRequest in boricaTransaction.PaymentRequests)
                                        {
                                            if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess)
                                            {
                                                paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Canceled;
                                                paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                                                if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                                                {
                                                    Notification statusNotification = new Notification(paymentRequest, (int?)null);

                                                    unitOfWork.DbContext.Set<Notification>().Add(statusNotification);
                                                }
                                            }

                                        }

                                        apgwResponse.UpdateBoricaTransactionCanceledPayment(boricaTransaction);
                                        unitOfWork.Save();
                                    }
                                    else if (boricaTransaction.TransactionDate.AddMinutes(20) <= DateTime.Now)
                                    {
                                        foreach (var paymentRequest in boricaTransaction.PaymentRequests)
                                        {
                                            if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess)
                                            {
                                                paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Pending;
                                                paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                                                if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                                                {
                                                    Notification statusNotification = new Notification(paymentRequest, (int?)null);

                                                    unitOfWork.DbContext.Set<Notification>().Add(statusNotification);
                                                }
                                            }

                                        }

                                        apgwResponse.UpdateBoricaTransactionCanceledPayment(boricaTransaction);
                                        unitOfWork.Save();
                                    }
                                }
                                else
                                {
                                    JobLogger.Get(JobName.BoricaUnprocessedRequestsJob).Log(LogLevel.Warn,
                                        string.Format("Грешка при опит на прочитане на данните от Борика. Моля проверете сертификата. Транзакция - {0}, дата - {1}",
                                        boricaTransaction.Order,
                                        DateTime.Now.ToLongDateString()));
                                }
                            }

                            if (boricaTransaction.TransactionStatusId == pendingPayment &&
                                boricaTransaction.TransactionDate.AddHours(finalLimit) < transactionStartDate)
                            {
                                boricaTransaction.TransactionStatusId = cancelPayment;
                                unitOfWork.Save();
                            }
                        }
                        catch (Exception ex)
                        {
                            JobLogger.Get(JobName.BoricaUnprocessedRequestsJob).Log(LogLevel.Warn, ex.Message, ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.BoricaUnprocessedRequestsJob).Log(LogLevel.Warn, ex.Message, ex);
            }
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                DependencyFactory = null;
                _isDisposed = !_isDisposed;
            }
        }

        private void UpdateNetProtocol()
        {
            int securityProtocol = (int)(SecurityProtocolType.Tls12 |
                SecurityProtocolType.Tls11 |
                SecurityProtocolType.Tls);

            if (((int)ServicePointManager.SecurityProtocol & securityProtocol) != securityProtocol)
            {
                ServicePointManager.SecurityProtocol |=
                    SecurityProtocolType.Tls12 |
                    SecurityProtocolType.Tls11 |
                    SecurityProtocolType.Tls;
            }
        }

        private APGWStatusCheckRequestDataVO CreateBoricaRequestData(BoricaTransaction boricaTransaction)
        {
            int orderId = BoricaCvposHelper.GetTransactionOrder(boricaTransaction.Order);

            APGWStatusCheckRequestDataVO apgwRequest = new APGWStatusCheckRequestDataVO()
            {
                Terminal = AppSettings.EPaymentsWeb_CentralVposDevTerminalId,
                Order = orderId.ToString("000000"),
                TRAN_TRTYPE = TrTypeEnum.Payment,
                Date = new DateTime(boricaTransaction.TransactionDate.Year, boricaTransaction.TransactionDate.Month, boricaTransaction.TransactionDate.Day),
            };

            using (var certificate = new X509Certificate2(BoricaCvposHelper.GetBoricaCvposDevPfxPath(),
                        AppSettings.EPaymentsWeb_CentralVposPrivateKeyPassphrase,
                        X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable))
            {
                using (RSA rsa = certificate.GetRSAPrivateKey())
                {
                    byte[] pSignBytes = Encoding.UTF8.GetBytes(apgwRequest.GetPSignData());
                    byte[] signature = rsa.SignData(pSignBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                    var isValid = rsa.VerifyData(pSignBytes, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                    if (isValid == false)
                    {
                        throw new Exception("Грешка при валидация на подписаните на данни при изпращане към ЦВПОС.");
                    }

                    apgwRequest.P_Sign = BinHexHelper.ByteArrayToHexString(signature).ToUpper();
                }
            }

            return apgwRequest;
        }

        private bool VerifyResponse(APGWPaymentResponseDataDO apgwResponse)
        {
            using (var certificate = BoricaCvposHelper.GetBoricaCvposDev())
            {
                if(certificate == null)
                {
                    JobLogger.Get(JobName.BoricaUnprocessedRequestsJob).Log(LogLevel.Warn, $"apgwResponse.Order: {apgwResponse.Order}. The certificate is null.");
                }

                using (var rsa = certificate.GetRSAPublicKey())
                {
                    if(rsa == null)
                    {
                        JobLogger.Get(JobName.BoricaUnprocessedRequestsJob).Log(LogLevel.Warn, $"apgwResponse.Order: {apgwResponse.Order}. The RSAPublicKey is null.");
                    }

                    byte[] data = Encoding.UTF8.GetBytes(apgwResponse.GetPSignData());
                    byte[] signature = BinHexHelper.HexStringToByteArray(apgwResponse.P_Sign);

                    return rsa.VerifyData(data, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                }
            }
        }
    }
}