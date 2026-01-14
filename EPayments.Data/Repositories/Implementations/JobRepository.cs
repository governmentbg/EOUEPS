using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using System.Collections.Generic;
using System.Linq;
using EPayments.Model.Models;
using EPayments.Model.Enums;
using EPayments.Common.Helpers;
using EPayments.Data.ViewObjects.Api;
using System;
using System.Data.Entity;
//using EPayments.Data.ViewObjects.Web;

namespace EPayments.Data.Repositories.Implementations
{
    internal class JobRepository : BaseRepository, IJobRepository
    {
        public JobRepository(IUnitOfWork unitOfWork)
            : base(unitOfWork)
        {
        }

        public IList<int> GetPendingEmailIds(int limit, int maxFailedAttempts, TimeSpan failedAttemptTimeout)
        {
            var maxInterval = DateTime.Now - failedAttemptTimeout;

            return this.unitOfWork.DbContext.Set<Email>()
                .Where(e => e.NotificationStatusId == NotificationStatus.Pending && (e.FailedAttempts == 0 || (e.FailedAttempts < maxFailedAttempts && e.ModifyDate < maxInterval)))
                .OrderBy(e => e.CreateDate)
                .Select(e => e.EmailId)
                .Take(limit)
                .ToList();
        }

        public Email GetEmailById(int emailId)
        {
            return this.unitOfWork.DbContext.Set<Email>().SingleOrDefault(e => e.EmailId == emailId);
        }

        public ICollection<int> GetPendingDeliveryEserviceNotificationIds(int limit)
        {
            var lefts = this.unitOfWork.DbContext.Set<PaymentRequestObligationLog>()
                    .Where(prol => !prol.EserviceDeliveryNotifications.Any())
                    .Take(limit)
                    .Select(prol => prol.PaymentRequestObligationLogsId)
                    .ToList();

            var rights = this.unitOfWork.DbContext.Set<EserviceDeliveryNotification>()
                .Where(esdn => (esdn.NotificationStatusId == NotificationStatus.Pending || esdn.NotificationStatusId == NotificationStatus.Error) && (esdn.SendNotBefore == null || esdn.SendNotBefore <= DateTime.Now))
                .Take(limit)
                .Select(esdn => esdn.PaymentRequestObligationLogsId)
                .ToList();

            HashSet<int> cross = new HashSet<int>(lefts);
            cross.UnionWith(rights);

            return cross;
        }

        public IList<int> GetPendingEserviceNotificationIds(int limit)
        {
           return this.unitOfWork.DbContext.Set<EserviceNotification>()
                .Where(e => e.NotificationStatusId == NotificationStatus.Pending && (e.SendNotBefore == null || e.SendNotBefore <= DateTime.Now))
                .OrderBy(e => e.CreateDate)
                .Select(e => e.EserviceNotificationId)
                .Take(limit)
                .ToList();
        }

        public EserviceNotification GetEserviceNotificationById(int eserviceNotificationId)
        {
            return this.unitOfWork.DbContext.Set<EserviceNotification>()
                                .Include(e => e.EserviceClient)
                                .Include(e => e.PaymentRequest)
                                .Include(e => e.PaymentRequest.ObligationType)
                                .SingleOrDefault(e => e.EserviceNotificationId == eserviceNotificationId);
        }

        public PaymentRequestObligationLog GetDeliveryEserviceNotificationById(int paymentRequestObligationLogId)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequestObligationLog>()
                .Include(e => e.PaymentRequest)
                .Include(e => e.PaymentRequest.EserviceClient)
                .Include(e => e.PaymentRequest.ObligationType)
                .SingleOrDefault(e => e.PaymentRequestObligationLogsId == paymentRequestObligationLogId);
        }

        public EserviceDeliveryNotification GetDeliveryEserviceNotificationByLogId(int paymentRequestObligationLogId)
        {
            return this.unitOfWork.DbContext.Set<EserviceDeliveryNotification>()
                .Include(e => e.PaymentRequestObligationLog)
                .SingleOrDefault(e => e.PaymentRequestObligationLogsId == paymentRequestObligationLogId);
        }

  
        public List<BoricaTransaction> GetPendingCVposTransactions(BoricaTransactionStatusEnum boricaTransactionStatus,int batchSize)
        {
            return this.unitOfWork.DbContext.Set<BoricaTransaction>()
                .Where(e => e.TransactionStatusId <= (int)BoricaTransactionStatusEnum.Pending)
                .OrderBy(e => e.BoricaTransactionId)
                .Take(batchSize)
                .ToList();
        }
        public IList<int> GetExpiredRequestIds()
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Pending && e.ExpirationDate <= DateTime.Now)
                .Select(e => e.PaymentRequestId)
                .ToList();
        }

        public PaymentRequest GetRequestById(int paymentRequestId)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Include(pr => pr.BoricaTransactions)
                .Include(pr => pr.ObligationType)
                .SingleOrDefault(e => e.PaymentRequestId == paymentRequestId);
        }

        public IList<EserviceClient> GetAllActiveEserviceClients()
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .Where(e => e.IsActive)
                .ToList();
        }

        public IList<int> GetPendingEventRegisterNotificationIds()
        {
            return this.unitOfWork.DbContext.Set<EventRegisterNotification>()
                .Where(e => e.NotificationStatusId == NotificationStatus.Pending && (e.SendNotBefore == null || e.SendNotBefore <= DateTime.Now))
                .OrderBy(e => e.CreateDate)
                .Select(e => e.EventRegisterNotificationId)
                .ToList();
        }

        public EventRegisterNotification GetEventRegisterNotificationById(int eventRegisterNotificationId)
        {
            return this.unitOfWork.DbContext.Set<EventRegisterNotification>()
                .SingleOrDefault(e => e.EventRegisterNotificationId == eventRegisterNotificationId);
        }

        public List<EserviceBankAccount> GetActiveEserviceBankAccountsWithUploadTransactions()
        {
            return this.unitOfWork.DbContext.Set<EserviceBankAccount>()
                .Where(e => e.UploadTransactions && e.IsActive)
                .ToList();
        }

        public List<PaymentRequest> GetPendingPaymentRequestsByIban(string iban)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Include(p => p.ObligationType)
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Pending && e.ServiceProviderIBAN == iban)
                .ToList();
        }

        public VposFiBankRequest GetPendingResultVposFiBankRequest(int maxFailedAttempts, TimeSpan failedAttemptTimeout, int fibankRequestTimeoutInMinutes)
        {
            var checkRequestsCreatedBefore = DateTime.Now.AddMinutes(-1 * (fibankRequestTimeoutInMinutes + 1));
            var checkRequestsWithLastCheckResultDateBefore = DateTime.Now - failedAttemptTimeout;

            return this.unitOfWork.DbContext.Set<VposFiBankRequest>()
                .Include(e => e.PaymentRequest.EserviceClient)
                .Include(e => e.PaymentRequest.ObligationType)
                .Where(e =>
                     e.ResultStatus == VposRequestResultStatus.Pending &&
                     e.CreateDate < checkRequestsCreatedBefore &&
                     e.JobCheckResultFailedAttempts < maxFailedAttempts &&
                     (e.JobLastCheckResultDate == null || e.JobLastCheckResultDate < checkRequestsWithLastCheckResultDateBefore))
                .OrderBy(e => e.VposFiBankRequestId)
                .FirstOrDefault();
        }

        public VposDskEcommRequest GetPendingResultVposDskEcommRequest(int maxFailedAttempts, TimeSpan failedAttemptTimeout, int fibankRequestTimeoutInMinutes)
        {
            var checkRequestsCreatedBefore = DateTime.Now.AddMinutes(-1 * (fibankRequestTimeoutInMinutes + 1));
            var checkRequestsWithLastCheckResultDateBefore = DateTime.Now - failedAttemptTimeout;

            return this.unitOfWork.DbContext.Set<VposDskEcommRequest>()
                .Include(e => e.PaymentRequest.EserviceClient.VposClient)
                .Include(e => e.PaymentRequest.ObligationType)
                .Where(e =>
                     e.ResultStatus == VposRequestResultStatus.Pending &&
                     e.CreateDate < checkRequestsCreatedBefore &&
                     e.JobCheckResultFailedAttempts < maxFailedAttempts &&
                     (e.JobLastCheckResultDate == null || e.JobLastCheckResultDate < checkRequestsWithLastCheckResultDateBefore))
                .OrderBy(e => e.VposFiBankRequestId)
                .FirstOrDefault();
        }

        public List<DistributionRevenue> GetDistributionsWithNoFiles(int limit, int skip)
        {
            return this.unitOfWork.DbContext
                        .Set<DistributionRevenue>()
                        .Where(dr => dr.IsFileGenerated == false && !dr.DistributionErrors.Any())
                        .OrderByDescending(dr => dr.DistributionRevenueId)
                        .Take(limit)
                        .Skip(skip)
                        .ToList();
        }

        public List<DistributionRevenuePayment> GetDistributionRevenuePayments(int distributionId)
        {
            return this.unitOfWork.DbContext
                        .Set<DistributionRevenuePayment>()
                        .Include(drp => drp.PaymentRequest)
                        .Include(drp => drp.EserviceClient.Department)
                        .Include(drp => drp.BoricaTransaction)
                        .Where(drp => drp.DistributionRevenueId == distributionId)
                        .ToList();
        }

        public bool HasDistributionForToday()
        {
            DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0, 0);
            DateTime endDate = startDate.AddMilliseconds(24 * 60 * 60 * 1000);

            return this.unitOfWork.DbContext
                .Set<DistributionRevenue>()
                .Any(dr => startDate <= dr.CreatedAt && dr.CreatedAt < endDate);
        }

        public PaymentRequest GetPaymentRequestById(int id)
        {
            return this.unitOfWork.DbContext
                .Set<PaymentRequest>()
                .FirstOrDefault(pr => pr.PaymentRequestId == id);
        }

        public void RemoveEserviceNotification(EserviceNotification notification)
        {
            this.unitOfWork.DbContext
                .Set<EserviceNotification>().Remove(notification);

            unitOfWork.Save();
        }

        public void RemoveEDelivaryNotification(PaymentRequestObligationLog notificationlog)
        {
            this.unitOfWork.DbContext
                              .Set<PaymentRequestObligationLog>()
                              .Remove(notificationlog);

            unitOfWork.Save();
        }
    }
}
