using Autofac.Features.OwnedInstances;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.CVPosTransaction.CVPosService;
using EPayments.CVPosTransaction.Manager;
using EPayments.Data.Repositories.Interfaces;
using EPayments.EventRegister.EventRegisterService;
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

namespace EPayments.Job.Host.Jobs.CVPosTransaction
{
    public class CVPosTransactionJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, ICVPosRegisterManager>>> dependencyFactory;
        private bool disposed;
        private TimeSpan period;
        private int batchSize;

        private readonly int cVposTransaction_MaxFailedAttempts =
            AppSettings.EPaymentsJobHost_CVPosTransactionJobMaxFailedAttempts;

        private readonly TimeSpan cVposTransaction_FailedAttemptTimeout =
            TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_CVPosTransactionJobFailedAttemptTimeoutInMinutes);

        private readonly int cVposTransaction_TimeoutInMinutes =
            AppSettings.EPaymentsJobHost_CVPosTransactionJobTimeoutInMinutes;

        private readonly string cVposTransaction_Agency =
            AppSettings.EPaymentsJobHost_CVPosTransactionJobAgency;

        private readonly string cVposTransaction_Event =
            AppSettings.EPaymentsJobHost_CVPosTransactionJobEvent;


        public CVPosTransactionJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository, ICVPosRegisterManager>>> dependencyFactory)
        {
            this.dependencyFactory = dependencyFactory;
            this.disposed = false;
            this.batchSize = AppSettings.PaymentsJobHost_CVPosTransactionJobBatchSize;
            this.period = TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_CVPosTransactionJobPeriodInMinutes);
        }

        public string Name
        {
            get { return "CVPosTransactionJob"; }
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
                
                int[] timeParams = AppSettings.CVPosTransactionJobStartTime.Split(new char[] { ':' }).Select(str => int.Parse(str)).ToArray();

                DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, timeParams[0], timeParams[1], 0, 0);

                if (true/*DateTime.Now > startDate*/)
                {
                    GlobalValue lastInvocationTime = systemRepository.GetGlobalValueByKey(GlobalValueKey.CVPosTransactionJobInvocationTime);

                    if (lastInvocationTime.ModifyDate.Year == DateTime.Now.Year &&
                        lastInvocationTime.ModifyDate.Month == DateTime.Now.Month &&
                        lastInvocationTime.ModifyDate.Day == DateTime.Now.Day)
                    {
                        return;
                    }

                    lastInvocationTime.ModifyDate = DateTime.Now;

                    unitOfWork.Save();

                    try
                    {
                        if (disposed)
                        {
                            return;
                        }

                        ProcessPendingCVposRequest(ct);
                    }
                    catch (OperationCanceledException ex)
                    {
                        JobLogger.Get(JobName.CVPosTransaction)
                            .Log(LogLevel.Error, "Job was canceled due to a token cancellation request", ex);
                    }
                    catch (Exception ex)
                    {
                        JobLogger.Get(JobName.CVPosTransaction)
                            .Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
                    }
                }
            }
        }

        public void ProcessPendingCVposRequest(CancellationToken ct)
        {
            DateTime jobDate = DateTime.Now;
            using (var factory = this.dependencyFactory())
            {
                var unitOfWork = factory.Value.Item1;
                var jobRepository = factory.Value.Item2;
                var systemRepository = factory.Value.Item3;
                var CVPosRegisterManager = factory.Value.Item4;
                int requestTimeout = AppSettings.EPaymentsJobHost_CVPosTransactionJobTimeoutBetweenRequestsInMilliseconds;
                string prefix = AppSettings.EPaymentsWeb_CentralVposPrefixHelper;
                bool testMode = AppSettings.EPaymentsJobHost_CVPosTransactionJobTestMode;
                if (ct.IsCancellationRequested)
                {
                    return;
                }

                Exception exception = null;

                DateTime startDate = DateTime.Now;

                BoricaTransaction oldestBoricaTransaction = unitOfWork.DbContext.Set<BoricaTransaction>()
                    .Where(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid &&
                        (bt.JobCheckResultFailedAttempts == null || bt.JobCheckResultFailedAttempts < this.cVposTransaction_MaxFailedAttempts))
                    .OrderBy(bt => bt.TransactionDate)
                    .FirstOrDefault();

                if (oldestBoricaTransaction != null && oldestBoricaTransaction.TransactionDate < startDate)
                {
                    startDate = oldestBoricaTransaction.TransactionDate;
                }

                startDate = new DateTime(startDate.Year, startDate.Month, startDate.Day);

                // for some reason borica service throw exception on first connection
                bool hasFirstTimeError = false;

                while (startDate < DateTime.Now)
                {
                    try
                    {
                        if (testMode) 
                        {
                            JobLogger.Get(JobName.CVPosTransaction)
                              .Log(LogLevel.Info, $"Активиран е тестови режим на работа {DateTime.Now}");
                            TestModeFaker(unitOfWork, ct);
                            return;
                        }

                        JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"ProcessPendingCVposRequest -> cVposTransaction_Agency:{cVposTransaction_Agency}, " +
                            $"cVposTransaction_Event:{cVposTransaction_Event}, startDate:{startDate}, AppSettings.EPaymentsWeb_CentralVposDevTerminalId:{AppSettings.EPaymentsWeb_CentralVposDevTerminalId}");

                        //recEventTransaction[] result = BoricaRetryPolicy.GetBoricaRetryPolicy(JobName.CVPosTransaction).Execute(() =>
                        //{
                        //   return CVPosRegisterManager.GetTransactionPerDate(
                        //       cVposTransaction_Agency,
                        //       cVposTransaction_Event,
                        //       startDate,
                        //       AppSettings.EPaymentsWeb_CentralVposDevTerminalId);
                        //});

                        recEventTransaction[] result = new recEventTransaction[2] ;
                        result[0] = new recEventTransaction()
                        {
                            rrn = "600000000001",
                            orderId = prefix + "251104182267-a1",
                            tax = 1.00M,
                            taxBorica = 1.00M,
                            settlementDate = DateTime.Now,
                            authorizationCode = "B85418",
                            tid = "TERM001",
                            distributionDate = DateTime.Now,
                            productCategory = "CAT001",
                            areaOfIssue = "AREA001",
                            cardBrand = "VISA"
                        };
                        result[1] = new recEventTransaction()
                        {
                            rrn = "600000000002",
                            orderId = prefix + "251104182267-a2",
                            tax = 1.00M,
                            taxBorica = 1.00M,
                            settlementDate = DateTime.Now,
                            authorizationCode = "B85418",
                            tid = "TERM001",
                            distributionDate = DateTime.Now,
                            productCategory = "CAT001",
                            areaOfIssue = "AREA001",
                            cardBrand = "VISA"
                        };

                        if (result == null || result.Length == 0)
                        {
                            JobLogger.Get(JobName.CVPosTransaction)
                                .Log(LogLevel.Info, String.Format("Няма данни за минал сетълмент на дата {0}. Проверката е направена на {1}", startDate, DateTime.Now));
                            DateTime sDate = new DateTime(startDate.Year, startDate.Month, startDate.Day, 0, 0, 0, 0);
                            DateTime eDate = sDate.AddDays(1);

                            List<BoricaTransaction> boricaTransactions = unitOfWork.DbContext.Set<BoricaTransaction>()
                                .Where(bt => (sDate <= bt.TransactionDate && bt.TransactionDate < eDate) && 
                                    bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid)
                                .ToList();

                            if (boricaTransactions != null && boricaTransactions.Count > 0)
                            {
                                using (DbContextTransaction transaction = unitOfWork.BeginTransaction())
                                {
                                    boricaTransactions.ForEach(tr =>
                                    {
                                        if (tr.JobCheckResultFailedAttempts == null)
                                        {
                                            tr.JobCheckResultFailedAttempts = 0;
                                        }

                                        tr.JobCheckResultFailedAttempts += 1;
                                        tr.JobLastCheckResultDate = DateTime.Now;
                                    });

                                    unitOfWork.Save();

                                    transaction.Commit();
                                }
                            }
                        }
                        else
                        {
                            foreach (recEventTransaction recEventTransaction in result)
                            {
                                string transactionOrderNumbersPart = recEventTransaction.orderId.Replace(prefix, string.Empty).TrimStart(new char[] { '0' });

                                BoricaTransaction currentBoricaTransaction = unitOfWork.DbContext.Set<BoricaTransaction>()
                                    .Include(bt => bt.PaymentRequests.Select(pr => pr.ObligationType))                                                                  
                                    .FirstOrDefault(bt => bt.Rrn == recEventTransaction.rrn && bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid);

                                JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"Запитване за дата: {startDate}, Отговор : borica transaction id {currentBoricaTransaction?.BoricaTransactionId} response: {JsonConvert.SerializeObject(recEventTransaction)}");
                                
                                if (currentBoricaTransaction != null)
                                {
                                    using (DbContextTransaction transaction = unitOfWork.BeginTransaction())
                                    {
                                        currentBoricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.TaxReceived;
  
                                        //Check if still old transactions in bgn 
                                        //if (currentBoricaTransaction.TransactionDate.Year <= 2025 && CurrencyHelper.IsEuroTimePeriod)
                                        //{
                                        //    currentBoricaTransaction.Fee = recEventTransaction.tax * CurrencyHelper.BgnToEuroRate;
                                        //    currentBoricaTransaction.Commission = recEventTransaction.taxBorica  * CurrencyHelper.BgnToEuroRate;
                                        //}
                                        //else
                                        //{
                                            currentBoricaTransaction.Fee = recEventTransaction.tax;
                                            currentBoricaTransaction.Commission = recEventTransaction.taxBorica;
                                        //}
                                        currentBoricaTransaction.SettlementDate = recEventTransaction.settlementDate;
                                        currentBoricaTransaction.AuthorizationCode = recEventTransaction.authorizationCode;
                                        currentBoricaTransaction.Terminal = recEventTransaction.tid;
                                        currentBoricaTransaction.DistributionDate = recEventTransaction.distributionDate;
                                        currentBoricaTransaction.ProductCategory = recEventTransaction.productCategory;
                                        currentBoricaTransaction.AreaOfIssue = recEventTransaction.areaOfIssue;
                                        currentBoricaTransaction.Card = recEventTransaction.cardBrand;
                                        currentBoricaTransaction.BoricaTransactionSettlement = JsonConvert.SerializeObject(recEventTransaction);

                                        //foreach (PaymentRequest paymentRequest in currentBoricaTransaction.PaymentRequests)
                                        //{
                                        //    if (paymentRequest.ObligationStatusId != ObligationStatusEnum.ForDistribution
                                        //        || paymentRequest.ObligationStatusId != ObligationStatusEnum.CheckedAccount) 
                                        //    {
                                        //        paymentRequest.ObligationStatusId = ObligationStatusEnum.Paid;
                                        //    }
                                            
                                        //    if (currentBoricaTransaction.Rc == "00" 
                                        //        && currentBoricaTransaction.Action == "0"
                                        //        && (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess || paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending))
                                        //    {
                                        //        paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Paid;
                                        //        paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                                        //        string[] applicantUins = currentBoricaTransaction.PaymentRequests
                                        //                                     .Where(pr => !string.IsNullOrWhiteSpace(pr.ApplicantUin))
                                        //                                     .Select(pr => pr.ApplicantUin)
                                        //                                     .ToArray();
                                        //        User[] users = unitOfWork.DbContext.Set<User>()
                                        //                            .Where(u => applicantUins.Contains(u.Egn))
                                        //                            .ToArray();

                                        //        User user = users.FirstOrDefault(u => string.Equals(u.Egn, paymentRequest.ApplicantUin, StringComparison.OrdinalIgnoreCase));

                                        //        if (user != null && !String.IsNullOrWhiteSpace(user.Email))
                                        //        {
                                        //            if (user.StatusObligationNotifications)
                                        //            {
                                        //                unitOfWork.DbContext.Set<Mail>().Add(user.CreateStatusObligationNotificationEmail(paymentRequest));
                                        //            }

                                        //            if (user.StatusNotifications)
                                        //            {
                                        //                unitOfWork.DbContext.Set<Mail>().Add(user.CreateStatusNotificationEmail(paymentRequest));
                                        //            }
                                        //        }

                                        //        if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                                        //        {
                                        //            Notification statusNotification = new Notification(paymentRequest, (int?)null);

                                        //            unitOfWork.DbContext.Set<Notification>().Add(statusNotification);
                                        //        }
                                        //    }
                                        //}

                                        unitOfWork.Save();
                                        transaction.Commit();
                                    }
                                }

                                if (ct.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                        }

                        startDate = startDate.AddDays(1);

                        if (hasFirstTimeError == false)
                        {
                            hasFirstTimeError = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (hasFirstTimeError == false)
                        {
                            hasFirstTimeError = true;
                        }
                        else
                        {
                            exception = ex;
                            int typeEx = 0;

                            if (ex.Data.Contains("Type"))
                            {
                                typeEx = (int)ex.Data["Type"];
                            }

                            switch (typeEx)
                            {
                                case (int)CVPosExeptionEnums.Certificate:
                                    JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"{ex.Message}, StackTrace -> {ex.StackTrace}");
                                    break;
                                case (int)CVPosExeptionEnums.Communication:
                                    JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"{ex.Message}, StackTrace -> {ex.StackTrace}");
                                    break;
                                case (int)CVPosExeptionEnums.EndpointNotFound:
                                    JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"{ex.Message}, StackTrace -> {ex.StackTrace}");
                                    break;
                                case (int)CVPosExeptionEnums.Fault:
                                    {
                                        var code = ex.Data["Code"].ToString();
                                        var msg = ex.Data["Message"].ToString();
                                        JobLogger.Get(JobName.CVPosTransaction)
                                            .Log(LogLevel.Info, String.Format("{0} FaultCode: {1} Message:{2}", $"{ex.Message}, StackTrace -> {ex.StackTrace}", code, msg));
                                    }
                                    break;
                                default:
                                    JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"{ex.Message}, StackTrace -> {ex.StackTrace}");
                                    break;
                            }

                            throw;
                        }
                    }

                    Thread.Sleep(requestTimeout);
                }

                unitOfWork.Save();
            }
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;

                JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, "CVPosTransaction job disposed");
            }
        }

        private void TestModeFaker(IUnitOfWork unitOfWork, CancellationToken ct) 
        {
            var rnd = new Random();
            var fee = (decimal)rnd.NextDouble();
            var commission = (decimal)rnd.NextDouble();
            var authorizationCode = rnd.Next(1, 1000).ToString();
            var terminal = rnd.Next(1001, 2000).ToString();
            var productCategory = rnd.Next(2001, 3000).ToString();
            var areaOfIssue = rnd.Next(3001, 4000).ToString();
            var card = rnd.Next(4001, 5000).ToString();

            List<BoricaTransaction> listBoricaTransaction = unitOfWork.DbContext.Set<BoricaTransaction>()
                                    .Include(bt => bt.PaymentRequests)
                                    .Where(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid)
                                    .ToList();

            foreach (var currentBoricaTransaction in listBoricaTransaction) 
            {
                if (currentBoricaTransaction != null)
                {
                    using (DbContextTransaction transaction = unitOfWork.BeginTransaction())
                    {
                        currentBoricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.TaxReceived;
                        currentBoricaTransaction.Fee = fee;
                        currentBoricaTransaction.Commission = commission;
                        currentBoricaTransaction.SettlementDate = DateTime.Now;
                        currentBoricaTransaction.AuthorizationCode = authorizationCode;
                        currentBoricaTransaction.Terminal = terminal;
                        currentBoricaTransaction.DistributionDate = DateTime.Now;
                        currentBoricaTransaction.ProductCategory = productCategory;
                        currentBoricaTransaction.AreaOfIssue = areaOfIssue;
                        currentBoricaTransaction.Card = card;

                        foreach (PaymentRequest paymentRequest in currentBoricaTransaction.PaymentRequests)
                        {
                            paymentRequest.ObligationStatusId = ObligationStatusEnum.Paid;
                        }

                        JobLogger.Get(JobName.CVPosTransaction).Log(LogLevel.Info, $"Таблицата BoricaTransactions с Id={currentBoricaTransaction.BoricaTransactionId} транзакции беше обновена с автоматично генерирани данни Fee={fee}, Commission={commission}");

                        unitOfWork.Save();

                        transaction.Commit();
                    }
                }

                if (ct.IsCancellationRequested)
                {
                    return;
                }
            }
        }
    }
}
