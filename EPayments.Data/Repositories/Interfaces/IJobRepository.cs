using EPayments.Data.ViewObjects.Api;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IJobRepository : IBaseRepository
    {
        IList<int> GetPendingEmailIds(int limit, int maxFailedAttempts, TimeSpan failedAttemptTimeout);

        Email GetEmailById(int emailId);

        ICollection<int> GetPendingDeliveryEserviceNotificationIds(int limit);

        IList<int> GetPendingEserviceNotificationIds(int limit);

        EserviceNotification GetEserviceNotificationById(int eserviceNotificationId);

        PaymentRequestObligationLog GetDeliveryEserviceNotificationById(int eserviceNotificationId);

        EserviceDeliveryNotification GetDeliveryEserviceNotificationByLogId(int paymentRequestObligationLogId);

        List<BoricaTransaction> GetPendingCVposTransactions(BoricaTransactionStatusEnum boricaTransactionStatus,int batchSize);

        IList<int> GetExpiredRequestIds();

        PaymentRequest GetRequestById(int paymentRequestId);

        IList<EserviceClient> GetAllActiveEserviceClients();

        IList<int> GetPendingEventRegisterNotificationIds();

        EventRegisterNotification GetEventRegisterNotificationById(int eventRegisterNotificationId);

        List<EserviceBankAccount> GetActiveEserviceBankAccountsWithUploadTransactions();

        List<PaymentRequest> GetPendingPaymentRequestsByIban(string iban);

        PaymentRequest GetPaymentRequestById(int id);

        VposFiBankRequest GetPendingResultVposFiBankRequest(int maxFailedAttempts, TimeSpan failedAttemptTimeout, int fibankRequestTimeoutInMinutes);

        VposDskEcommRequest GetPendingResultVposDskEcommRequest(int maxFailedAttempts, TimeSpan failedAttemptTimeout, int fibankRequestTimeoutInMinutes);

        List<DistributionRevenue> GetDistributionsWithNoFiles(int limit, int skip);

        List<DistributionRevenuePayment> GetDistributionRevenuePayments(int distributionId);

        bool HasDistributionForToday();

        void RemoveEserviceNotification(EserviceNotification notification);

        void RemoveEDelivaryNotification(PaymentRequestObligationLog notificationlog);
    }
}
