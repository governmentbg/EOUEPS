using EPayments.Common.Data;
using EPayments.Model.Models;
using System.Data.Entity;

namespace EPayments.Model
{
    public class ModelDbConfiguration : IDbConfiguration
    {
        public void AddConfiguration(DbModelBuilder modelBuilder)
        {
            modelBuilder.Configurations.Add(new BankMap());
            modelBuilder.Configurations.Add(new DepartmentMap());
            modelBuilder.Configurations.Add(new PaymentRequestMap());
            modelBuilder.Configurations.Add(new PaymentRequestStatusLogMap());
            modelBuilder.Configurations.Add(new PaymentRequestXmlMap());
            modelBuilder.Configurations.Add(new CertificateMap());
            modelBuilder.Configurations.Add(new UserMap());
            modelBuilder.Configurations.Add(new EmailMap());
            modelBuilder.Configurations.Add(new DisclaimerTemplateMap());
            modelBuilder.Configurations.Add(new DisclaimerMap());
            modelBuilder.Configurations.Add(new EbankingAccessLogMap());
            modelBuilder.Configurations.Add(new ErrorLogMap());
            modelBuilder.Configurations.Add(new LoginAttemptLogMap());
            modelBuilder.Configurations.Add(new PaymentRequestIdentifierMap());
            modelBuilder.Configurations.Add(new EserviceClientMap());
            modelBuilder.Configurations.Add(new VposClientMap());
            modelBuilder.Configurations.Add(new EbankingClientMap());
            modelBuilder.Configurations.Add(new VposResultMap());
            modelBuilder.Configurations.Add(new VposRedirectMap());
            modelBuilder.Configurations.Add(new VposBoricaRequestMap());
            modelBuilder.Configurations.Add(new VposFiBankRequestMap());
            modelBuilder.Configurations.Add(new VposDskEcommRequestMap());
            modelBuilder.Configurations.Add(new VposEpayRequestMap());
            modelBuilder.Configurations.Add(new AuthPassLoginMap());
            modelBuilder.Configurations.Add(new EserviceNotificationMap());
            modelBuilder.Configurations.Add(new EserviceDeliveryNotificationMap());
            modelBuilder.Configurations.Add(new EventRegisterNotificationMap());
            modelBuilder.Configurations.Add(new GlobalValueMap());
            modelBuilder.Configurations.Add(new TransactionFileMap());
            modelBuilder.Configurations.Add(new TransactionRecordMap());
            modelBuilder.Configurations.Add(new EserviceAdminUserMap());
            modelBuilder.Configurations.Add(new EserviceBankAccountMap());
            modelBuilder.Configurations.Add(new EserviceAdminUserBankAccountMap());
            modelBuilder.Configurations.Add(new InternalAdminUserMap());
            modelBuilder.Configurations.Add(new ObligationStatusMap());
            modelBuilder.Configurations.Add(new PaymentRequestHistoryMap());
            modelBuilder.Configurations.Add(new DistributionTypeMap());
            modelBuilder.Configurations.Add(new ObligationTypeMap());
            modelBuilder.Configurations.Add(new BoricaTransactionMap());
            modelBuilder.Configurations.Add(new DistributionRevenueMap());
            modelBuilder.Configurations.Add(new DistributionRevenuePaymentsMap());
            modelBuilder.Configurations.Add(new DistributionErrorMap());
            modelBuilder.Configurations.Add(new BoricaTransactionStatusMap());
            modelBuilder.Configurations.Add(new PaymentRequestStatusMap());
        }
    }
}