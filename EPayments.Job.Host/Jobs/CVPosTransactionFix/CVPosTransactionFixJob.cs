using Autofac.Features.OwnedInstances;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.CVPosTransaction.CVPosService;
using EPayments.CVPosTransaction.Manager;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Job.Host.Core;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading;
using Mail = EPayments.Model.Models.Email;
using Notification = EPayments.Model.Models.EserviceNotification;

namespace EPayments.Job.Host.Jobs.CVPosTransactionFix
{
    public class CVPosTransactionFixJob : IJob
    {
        private readonly Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, ICVPosRegisterManager>>> dependencyFactory;
        private bool disposed = false;
        private TimeSpan period = TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_CVPosTransactionFixJobPeriodInMinutes);
        private readonly int maxFailedAttempts = AppSettings.EPaymentsJobHost_CVPosTransactionFixJobMaxFailedAttempts;
        private readonly string agency = AppSettings.EPaymentsJobHost_CVPosTransactionFixJobAgency;
        private readonly string cVposTransaction_Event = AppSettings.EPaymentsJobHost_CVPosTransactionFixJobEvent;

        public CVPosTransactionFixJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, ICVPosRegisterManager>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
        }

        public string Name
        {
            get { return "CVPosTransactionFixJob"; }
        }

        public TimeSpan Period
        {
            get { return this.period; }
        }

        public void Action(CancellationToken cancellationToken)
        {
            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var systemRepository = factory.Value.Item3;

                int[] timeParams = AppSettings.EPaymentsJobHost_CVPosTransactionFixJobStartTime.Split(':').Select(str => int.Parse(str)).ToArray();

                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, timeParams[0], timeParams[1], 0, 0);

                if (DateTime.Now >= startDate)
                {
                    GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.CVPosTransactionFixJobInvocationTime);

                    if (lastInvocationTime.ModifyDate.Year == DateTime.Now.Year &&
                        lastInvocationTime.ModifyDate.Month == DateTime.Now.Month &&
                        lastInvocationTime.ModifyDate.Day == DateTime.Now.Day)
                    {
                        return;
                    }

                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();

                    if (disposed)
                    {
                        return;
                    }

                    try
                    {
                        ProcessPendingCVposFixRequest(cancellationToken);
                    }
                    catch (OperationCanceledException ex)
                    {
                        JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Error, "Job was canceled due to a token cancellation request", ex);
                    }
                    catch (Exception ex)
                    {
                        JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
                    }
                }
            }
        }

        public void ProcessPendingCVposFixRequest(CancellationToken cancellationToken)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            DateTime jobDate = DateTime.Now;
            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var jobRepository = factory.Value.Item2;
                var systemRepository = factory.Value.Item3;
                var CVPosRegisterManager = factory.Value.Item4;
                int requestTimeout = AppSettings.EPaymentsJobHost_CVPosTransactionFixJobTimeoutBetweenRequestsInMilliseconds;
                string prefix = AppSettings.EPaymentsWeb_CentralVposPrefixHelper;
                bool testMode = AppSettings.EPaymentsJobHost_CVPosTransactionFixJobTestMode;
                DateTime startDate = CreateDateTimeFromConfigString(AppSettings.EPaymentsJobHost_CVPosTransactionFixJobStartDate);
                DateTime endDate = CreateDateTimeFromConfigString(AppSettings.EPaymentsJobHost_CVPosTransactionFixJobEndDate);

                // for some reason borica service throw exception on first connection
                bool hasFirstTimeError = false;

                while (startDate < endDate)
                {
                    try
                    {
                        if (testMode)
                        {
                            JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, $"Активиран е тестови режим на работа {DateTime.Now}");
                            TestModeFaker(unitOfWork, cancellationToken);
                            return;
                        }

                        JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, $"ProcessPendingCVposFixRequest -> cVposTransaction_Agency:{agency}, " +
                            $"cVposTransaction_Event:{cVposTransaction_Event}, startDate:{startDate}, AppSettings.EPaymentsWeb_CentralVposDevTerminalId:{AppSettings.EPaymentsWeb_CentralVposDevTerminalId}");

                        recEventTransaction[] recEventTransactions = BoricaRetryPolicy.GetBoricaRetryPolicy(JobName.CVPosTransactionFix).Execute(() =>
                        {
                            return CVPosRegisterManager.GetTransactionPerDate(
                               agency,
                               cVposTransaction_Event,
                               startDate,
                               AppSettings.EPaymentsWeb_CentralVposDevTerminalId);
                        });
                       
                        if (recEventTransactions == null || recEventTransactions.Length == 0)
                        {
                            JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, String.Format("Няма данни за минал сетълмент на дата {0}. Проверката е направена на {1}", startDate, DateTime.Now));
                            DateTime sDate = new DateTime(startDate.Year, startDate.Month, startDate.Day, 0, 0, 0, 0);
                            DateTime eDate = sDate.AddDays(1);

                            //List<BoricaTransaction> boricaTransactions = unitOfWork.DbContext.Set<BoricaTransaction>()
                            //    .Where(bt => (sDate <= bt.TransactionDate && bt.TransactionDate < eDate) && bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Canceled)
                            //    .ToList();

                            //if (boricaTransactions != null && boricaTransactions.Count > 0)
                            //{
                            //    using (DbContextTransaction transaction = unitOfWork.BeginTransaction())
                            //    {
                            //        boricaTransactions.ForEach(tr =>
                            //        {
                            //            if (tr.JobCheckResultFailedAttempts == null)
                            //            {
                            //                tr.JobCheckResultFailedAttempts = 0;
                            //            }

                            //            tr.JobCheckResultFailedAttempts += 1;
                            //            tr.JobLastCheckResultDate = DateTime.Now;
                            //        });

                            //        unitOfWork.Save();
                            //        transaction.Commit();
                            //    }
                            //}
                        }
                        else
                        {
                            foreach (recEventTransaction recEventTransaction in recEventTransactions)
                            {
                                string transactionOrder = String.Format("{0}{1}", Formatter.DateToShortFormat(recEventTransaction.transactionDate),
                                                                        recEventTransaction.orderId.Replace(prefix, string.Empty).TrimStart(new char[] { '0' }));

                                BoricaTransaction currentBoricaTransaction = unitOfWork.DbContext.Set<BoricaTransaction>()
                                    .Include(bt => bt.PaymentRequests)
                                    .FirstOrDefault(bt => bt.Order == transactionOrder && bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Canceled);

                                JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, $"Запитване за дата: {startDate}, Отговор : borica transaction id {currentBoricaTransaction?.BoricaTransactionId} response: {JsonConvert.SerializeObject(recEventTransaction)}");

                                //if (currentBoricaTransaction != null)
                                //{
                                //    using (DbContextTransaction DbContextTransaction = unitOfWork.BeginTransaction())
                                //    {
                                //currentBoricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.TaxReceived;
                                //currentBoricaTransaction.Fee = recEventTransaction.tax;
                                //currentBoricaTransaction.Commission = recEventTransaction.taxBorica;
                                //currentBoricaTransaction.SettlementDate = recEventTransaction.settlementDate;
                                //currentBoricaTransaction.AuthorizationCode = recEventTransaction.authorizationCode;
                                //currentBoricaTransaction.Terminal = recEventTransaction.tid;
                                //currentBoricaTransaction.DistributionDate = recEventTransaction.distributionDate;
                                //currentBoricaTransaction.ProductCategory = recEventTransaction.productCategory;
                                //currentBoricaTransaction.AreaOfIssue = recEventTransaction.areaOfIssue;
                                //currentBoricaTransaction.Card = recEventTransaction.cardBrand;
                                //currentBoricaTransaction.BoricaTransactionSettlement = JsonConvert.SerializeObject(recEventTransaction);

                                //foreach (PaymentRequest paymentRequest in currentBoricaTransaction.PaymentRequests)
                                //{
                                //if (paymentRequest.ObligationStatusId != ObligationStatusEnum.ForDistribution
                                //    || paymentRequest.ObligationStatusId != ObligationStatusEnum.CheckedAccount)
                                //{
                                //    paymentRequest.ObligationStatusId = ObligationStatusEnum.Paid;
                                //}

                                //if (currentBoricaTransaction.Rc == "00"
                                //    && currentBoricaTransaction.Action == "0"
                                //    && (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess || paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending))
                                //{
                                //paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Paid;
                                //paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                                //string[] applicantUins = currentBoricaTransaction.PaymentRequests
                                //                             .Where(pr => !string.IsNullOrWhiteSpace(pr.ApplicantUin))
                                //                             .Select(pr => pr.ApplicantUin)
                                //                             .ToArray();
                                //User[] users = unitOfWork.DbContext.Set<User>().Where(u => applicantUins.Contains(u.Egn)).ToArray();

                                //User user = users.FirstOrDefault(u => string.Equals(u.Egn, paymentRequest.ApplicantUin, StringComparison.OrdinalIgnoreCase));

                                //if (user != null && !String.IsNullOrWhiteSpace(user.Email))
                                //{
                                //    if (user.StatusObligationNotifications)
                                //    {
                                //        unitOfWork.DbContext.Set<Mail>().Add(user.CreateStatusObligationNotificationEmail(paymentRequest));
                                //    }

                                //    if (user.StatusNotifications)
                                //    {
                                //        unitOfWork.DbContext.Set<Mail>().Add(user.CreateStatusNotificationEmail(paymentRequest));
                                //    }
                                //}

                                //if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                                //{
                                //    Notification statusNotification = new Notification(paymentRequest, (int?)null);

                                //    unitOfWork.DbContext.Set<Notification>().Add(statusNotification);
                                //}
                                //    }
                                //}

                                //unitOfWork.Save();
                                //DbContextTransaction.Commit();
                                //    }
                                //}

                                if (cancellationToken.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                        }

                        startDate = startDate.AddDays(1);
                        hasFirstTimeError = true;
                    }
                    catch (Exception ex)
                    {
                        if (hasFirstTimeError == false)
                        {
                            hasFirstTimeError = true;
                        }
                        else
                        {
                            int typeEx = 0;

                            if (ex.Data.Contains("Type"))
                            {
                                typeEx = (int)ex.Data["Type"];
                            }

                            switch (typeEx)
                            {
                                case (int)CVPosExeptionEnums.Fault:
                                    {
                                        var code = ex.Data["Code"].ToString();
                                        var msg = ex.Data["Message"].ToString();
                                        JobLogger.Get(JobName.CVPosTransactionFix)
                                            .Log(LogLevel.Info, String.Format("{0} FaultCode: {1} Message:{2}", $"{ex.Message}, StackTrace -> {ex.StackTrace}", code, msg));
                                    }
                                    break;
                                default:
                                    JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, $"{ex.Message}, StackTrace -> {ex.StackTrace}");
                                    break;
                            }

                            throw;
                        }
                    }

                    Thread.Sleep(requestTimeout);
                }

                //unitOfWork.Save();
            }
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;
                JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, "CVPosTransactionFix job disposed");
            }
        }

        private void TestModeFaker(IUnitOfWork unitOfWork, CancellationToken cancellationToken)
        {
            var random = new Random();
            var fee = (decimal)random.NextDouble();
            var commission = (decimal)random.NextDouble();
            var authorizationCode = random.Next(1, 1000).ToString();
            var terminal = random.Next(1001, 2000).ToString();
            var productCategory = random.Next(2001, 3000).ToString();
            var areaOfIssue = random.Next(3001, 4000).ToString();
            var card = random.Next(4001, 5000).ToString();

            List<BoricaTransaction> BoricaTransactions = unitOfWork.DbContext.Set<BoricaTransaction>()
                                    .Include(bt => bt.PaymentRequests)
                                    .Where(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Canceled)
                                    .ToList();

            foreach (var currentBoricaTransaction in BoricaTransactions)
            {
                if (currentBoricaTransaction != null)
                {
                    using (DbContextTransaction transaction = unitOfWork.BeginTransaction())
                    {
                        //currentBoricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.TaxReceived;
                        //currentBoricaTransaction.Fee = fee;
                        //currentBoricaTransaction.Commission = commission;
                        //currentBoricaTransaction.SettlementDate = DateTime.Now;
                        //currentBoricaTransaction.AuthorizationCode = authorizationCode;
                        //currentBoricaTransaction.Terminal = terminal;
                        //currentBoricaTransaction.DistributionDate = DateTime.Now;
                        //currentBoricaTransaction.ProductCategory = productCategory;
                        //currentBoricaTransaction.AreaOfIssue = areaOfIssue;
                        //currentBoricaTransaction.Card = card;

                        //foreach (PaymentRequest paymentRequest in currentBoricaTransaction.PaymentRequests)
                        //{
                        //    paymentRequest.ObligationStatusId = ObligationStatusEnum.Paid;
                        //}

                        JobLogger.Get(JobName.CVPosTransactionFix).Log(LogLevel.Info, $"Таблицата BoricaTransactions с Id={currentBoricaTransaction.BoricaTransactionId} транзакции беше обновена с автоматично генерирани данни Fee={fee}, Commission={commission}");

                        //unitOfWork.Save();
                        //transaction.Commit();
                    }
                }

                if (cancellationToken.IsCancellationRequested)
                {
                    return;
                }
            }
        }

        private DateTime CreateDateTimeFromConfigString(string configString)
        {
            int[] dateParts = configString.Split('.').Select(str => int.Parse(str)).ToArray();
            return new DateTime(dateParts[2], dateParts[1], dateParts[0]);
        }
    }
}
