using Autofac.Features.OwnedInstances;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using EPayments.Job.Host.Core;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using NLog;
using System.IO;
using EPayments.Job.Host.ProcessTransactionFiles.Parsers.UniCredit;
using EPayments.Model.Models;
using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using EPayments.Model.DataObjects.EmailTemplateContext;

namespace EPayments.Job.Host.EserviceNotification
{
    public class ProcessTransactionFilesJob : IJob
    {
        private Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> jobDependencyFactory;
        private object syncRoot = new object();
        private bool disposed;
        private TimeSpan period;

        public ProcessTransactionFilesJob(Func<Owned<DisposableTuple<IUnitOfWork, IJobRepository, ISystemRepository>>> jobDependencyFactory)
        {
            this.jobDependencyFactory = jobDependencyFactory;
            this.disposed = false;

            this.period = TimeSpan.FromSeconds(AppSettings.EPaymentsJobHost_ProcessTransactionFilesJobPeriodInSeconds);
        }

        public string Name
        {
            get { return "ProcessTransactionFilesJob"; }
        }

        public TimeSpan Period
        {
            get { return this.period; }
        }

        public void Action(CancellationToken ct)
        {
            try
            {
                List<EserviceBankAccount> eserviceBankAccounts = null;

                using (var factory = this.jobDependencyFactory())
                {
                    IJobRepository jobRepository = factory.Value.Item2;

                    eserviceBankAccounts = jobRepository.GetActiveEserviceBankAccountsWithUploadTransactions();
                }

                foreach (EserviceBankAccount bankAccount in eserviceBankAccounts)
                {
                    using (var factory_bankAccount = this.jobDependencyFactory())
                    {
                        IUnitOfWork unitOfWork_bankAccount = factory_bankAccount.Value.Item1;
                        IJobRepository jobRepository_bankAccount = factory_bankAccount.Value.Item2;
                        ISystemRepository systemRepository_bankAccount = factory_bankAccount.Value.Item3;

                        List<PendingPaymentRequest> pendingPaymentRequests =
                            jobRepository_bankAccount.GetPendingPaymentRequestsByIban(bankAccount.Iban)
                            .Select(e => new PendingPaymentRequest(e, false))
                            .ToList();

                        foreach (var unreadFile in Directory.GetFiles(bankAccount.TransactionsFilesPathUnread))
                        {
                            try
                            {
                                using (var factory_file = this.jobDependencyFactory())
                                {
                                    IUnitOfWork unitOfWork_file = factory_file.Value.Item1;
                                    IJobRepository jobRepository_file = factory_file.Value.Item2;

                                    UniCreditTransactionParser parser = new UniCreditTransactionParser(unreadFile);
                                    UniCreditTransactionFileDO transactionFileDO = parser.ExtractCreditTransactions();

                                    //TODO: check if UniCreditTransactionFileDO IBAN match expected IBAN

                                    if (transactionFileDO.Transactions.Count > 0)
                                    {
                                        TransactionFile newTransactionFile = new TransactionFile();
                                        newTransactionFile.EserviceBankAccountId = 1;
                                        newTransactionFile.FileName = unreadFile;
                                        newTransactionFile.BankStatementIban = Formatter.WhiteSpaceToNull(transactionFileDO.BankStatementIban);
                                        newTransactionFile.BankStatementDate = transactionFileDO.BankStatementDate;
                                        newTransactionFile.BankStatementNumber = Formatter.WhiteSpaceToNull(transactionFileDO.BankStatementNumber);
                                        newTransactionFile.CreateDate = DateTime.Now;

                                        foreach (UniCreditTransactionDO transactionDO in transactionFileDO.Transactions)
                                        {
                                            TransactionRecord newTransactionRecord = new TransactionRecord();
                                            newTransactionRecord.TransactionDate = transactionDO.TransactionDate;
                                            newTransactionRecord.TransactionAccountingDate = transactionDO.TransactionAccountingDate;
                                            newTransactionRecord.TransactionAmount = transactionDO.TransactionAmount;
                                            newTransactionRecord.TransactionReferenceId = Formatter.WhiteSpaceToNull(transactionDO.TransactionReferenceId);
                                            newTransactionRecord.InfoSystemTransactionType = Formatter.WhiteSpaceToNull(transactionDO.InfoSystemTransactionType);
                                            newTransactionRecord.InfoSystemTransactionDesc = Formatter.WhiteSpaceToNull(transactionDO.InfoSystemTransactionDesc);
                                            newTransactionRecord.InfoPaymentType = Formatter.WhiteSpaceToNull(transactionDO.InfoPaymentType);
                                            newTransactionRecord.InfoDocumentType = Formatter.WhiteSpaceToNull(transactionDO.InfoDocumentType);
                                            newTransactionRecord.InfoDocumentNumber = Formatter.WhiteSpaceToNull(transactionDO.InfoDocumentNumber);
                                            newTransactionRecord.InfoDocumentDate = transactionDO.InfoDocumentDate;
                                            if (newTransactionRecord.InfoDocumentNumber != null || newTransactionRecord.InfoDocumentDate.HasValue)
                                            {
                                                List<string> list = new List<string>();

                                                if (newTransactionRecord.InfoDocumentNumber != null)
                                                    list.Add(newTransactionRecord.InfoDocumentNumber);

                                                if (newTransactionRecord.InfoDocumentDate.HasValue)
                                                    list.Add(Formatter.DateToBgFormat(newTransactionRecord.InfoDocumentDate));

                                                newTransactionRecord.InfoDocumentNumberDate = String.Join(" / ", list);
                                            }
                                            newTransactionRecord.InfoPaymentPeriodBegining = transactionDO.InfoPaymentPeriodBegining;
                                            newTransactionRecord.InfoPaymentPeriodEnd = transactionDO.InfoPaymentPeriodEnd;
                                            newTransactionRecord.InfoDebtorBulstat = Formatter.WhiteSpaceToNull(transactionDO.InfoDebtorBulstat);
                                            newTransactionRecord.InfoDebtorEgn = Formatter.WhiteSpaceToNull(transactionDO.InfoDebtorEgn);
                                            newTransactionRecord.InfoDebtorLnch = Formatter.WhiteSpaceToNull(transactionDO.InfoDebtorLnch);
                                            newTransactionRecord.InfoDebtorName = Formatter.WhiteSpaceToNull(transactionDO.InfoDebtorName);

                                            if (newTransactionRecord.InfoDebtorBulstat != null || newTransactionRecord.InfoDebtorEgn != null || newTransactionRecord.InfoDebtorLnch != null)
                                            {
                                                List<string> list = new List<string>();

                                                if (newTransactionRecord.InfoDebtorBulstat != null)
                                                    list.Add(newTransactionRecord.InfoDebtorBulstat);

                                                if (newTransactionRecord.InfoDebtorEgn != null)
                                                    list.Add(newTransactionRecord.InfoDebtorEgn);

                                                if (newTransactionRecord.InfoDebtorLnch != null)
                                                    list.Add(newTransactionRecord.InfoDebtorLnch);

                                                newTransactionRecord.InfoDebtorBulstatEgnLnch = String.Join(" / ", list);
                                            }

                                            if (newTransactionRecord.InfoDebtorBulstatEgnLnch != null || newTransactionRecord.InfoDebtorName != null)
                                            {
                                                List<string> list = new List<string>();

                                                if (newTransactionRecord.InfoDebtorBulstatEgnLnch != null)
                                                    list.Add(newTransactionRecord.InfoDebtorBulstatEgnLnch);

                                                if (newTransactionRecord.InfoDebtorName != null)
                                                    list.Add(newTransactionRecord.InfoDebtorName);

                                                newTransactionRecord.InfoDebtorBulstatEgnLnchName = String.Join(" / ", list);
                                            }

                                            newTransactionRecord.InfoPaymentDetailsRaw = Formatter.WhiteSpaceToNull(transactionDO.InfoPaymentDetailsRaw);
                                            newTransactionRecord.InfoPaymentReason = Formatter.WhiteSpaceToNull(transactionDO.InfoPaymentReason);
                                            newTransactionRecord.InfoSenderBic = Formatter.WhiteSpaceToNull(transactionDO.InfoSenderBic);
                                            newTransactionRecord.InfoSenderIban = Formatter.WhiteSpaceToNull(transactionDO.InfoSenderIban);
                                            newTransactionRecord.InfoSenderName = Formatter.WhiteSpaceToNull(transactionDO.InfoSenderName);
                                            newTransactionRecord.InfoAC1AuthorizationCode = Formatter.WhiteSpaceToNull(transactionDO.InfoAC1AuthorizationCode);
                                            newTransactionRecord.InfoAC1BankCardInfo = Formatter.WhiteSpaceToNull(transactionDO.InfoAC1BankCardInfo);
                                            if (newTransactionRecord.InfoSenderName != null || newTransactionRecord.InfoSenderIban != null)
                                            {
                                                List<string> list = new List<string>();

                                                if (newTransactionRecord.InfoSenderIban != null)
                                                    list.Add(newTransactionRecord.InfoSenderIban);

                                                if (newTransactionRecord.InfoSenderName != null)
                                                    list.Add(newTransactionRecord.InfoSenderName);

                                                newTransactionRecord.InfoSenderIbanName = String.Join(" / ", list);
                                            }
                                            newTransactionRecord.TransactionRecordPaymentMethodId = transactionDO.InfoPaymentMethod;
                                            newTransactionRecord.TransactionRecordRefStatusId = TransactionRecordRefStatus.NotReferenced;

                                            newTransactionFile.TransactionRecords.Add(newTransactionRecord);
                                        }

                                        jobRepository_file.AddEntity<TransactionFile>(newTransactionFile);
                                        unitOfWork_file.Save();

                                        //bind transactions with TransactionRecordPaymentMethodId == TransactionRecordPaymentMethod.BankOrder

                                        foreach (TransactionRecord record in newTransactionFile.TransactionRecords.Where(e =>
                                            e.TransactionRecordPaymentMethodId == TransactionRecordPaymentMethod.BankOrder &&
                                            e.TransactionRecordRefStatusId == TransactionRecordRefStatus.NotReferenced))
                                        {
                                            if (pendingPaymentRequests.Where(e => !e.IsReferenced).Count() == 0)
                                                break;

                                            if (record.TransactionAmount.HasValue && record.InfoDocumentDate.HasValue && !String.IsNullOrWhiteSpace(record.InfoDocumentNumber))
                                            {
                                                DateTime docDate = record.InfoDocumentDate.Value;
                                                string docNumber = record.InfoDocumentNumber.Trim();

                                                PendingPaymentRequest referencedPendingPaymentRequest = pendingPaymentRequests.Where(e =>
                                                    !e.IsReferenced &&
                                                    e.PaymentRequest.CreateDate < newTransactionFile.CreateDate &&
                                                    e.PaymentRequest.PaymentReferenceDate == docDate &&
                                                    e.PaymentRequest.PaymentReferenceNumber == docNumber)
                                                    .FirstOrDefault();

                                                if (referencedPendingPaymentRequest != null)
                                                {
                                                    referencedPendingPaymentRequest.IsReferenced = true;

                                                    record.PaymentRequestId = referencedPendingPaymentRequest.PaymentRequest.PaymentRequestId;

                                                    if (record.TransactionAmount.Value > referencedPendingPaymentRequest.PaymentRequest.PaymentAmount)
                                                    {
                                                        record.TransactionRecordRefStatusId = TransactionRecordRefStatus.ReferencedOverpaidAmount;
                                                    }
                                                    else if (record.TransactionAmount.Value == referencedPendingPaymentRequest.PaymentRequest.PaymentAmount)
                                                    {
                                                        record.TransactionRecordRefStatusId = TransactionRecordRefStatus.ReferencedSuccessfully;
                                                    }
                                                    else
                                                    {
                                                        record.TransactionRecordRefStatusId = TransactionRecordRefStatus.ReferencedInsufficientAmount;
                                                    }

                                                    //set referenced payment request values

                                                    if (record.TransactionRecordRefStatusId == TransactionRecordRefStatus.ReferencedOverpaidAmount ||
                                                        record.TransactionRecordRefStatusId == TransactionRecordRefStatus.ReferencedSuccessfully)
                                                    {
                                                        referencedPendingPaymentRequest.PaymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Ordered;
                                                        referencedPendingPaymentRequest.PaymentRequest.PaidStatusPaymentMethodId = PaidStatusPaymentMethod.Other;
                                                        referencedPendingPaymentRequest.PaymentRequest.PaidStatusPaymentMethodDescription = "Платено по банков път";
                                                        referencedPendingPaymentRequest.PaymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;
                                                    }
                                                }
                                            }
                                        }

                                        unitOfWork_file.Save();
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                JobLogger.Get(JobName.ProcessTransactionFiles)
                                    .Log(LogLevel.Error, String.Format("Error occurred when parsing transaction file: {0}", unreadFile), ex);
                            }
                            finally
                            {
                                var fileName = Path.GetFileName(unreadFile);

                                foreach (var localFile in Directory.GetFiles(bankAccount.TransactionsFilesPathRead))
                                {
                                    if (Path.GetFileName(localFile).ToLower() == fileName.ToLower())
                                    {
                                        FileInfo lfi = new FileInfo(localFile);
                                        lfi.Delete();

                                        break;
                                    }
                                }

                                //move file
                                FileInfo fi = new FileInfo(unreadFile);
                                fi.MoveTo(bankAccount.TransactionsFilesPathRead + @"\" + fileName);
                            }
                        }

                        //send notifications for payment requests with changed status

                        foreach (PaymentRequest paymentRequest in
                            pendingPaymentRequests.Where(e => e.PaymentRequest.PaymentRequestStatusId != PaymentRequestStatus.Pending)
                            .Select(e => e.PaymentRequest))
                        {
                            User user = systemRepository_bankAccount.GetUserByUin(paymentRequest.ApplicantUin);
                            if (user != null && user.StatusNotifications && !String.IsNullOrWhiteSpace(user.Email))
                            {
                                StatusChangedPaymentRequestContextDO contextDO = new StatusChangedPaymentRequestContextDO(
                                    paymentRequest.PaymentRequestIdentifier,
                                    paymentRequest.ServiceProviderName,
                                    paymentRequest.PaymentReason,
                                    paymentRequest.PaymentAmount,
                                    paymentRequest.PaymentRequestStatusId.GetDescription(),
                                    paymentRequest.CreateDate);

                                Model.Models.Email email = new Model.Models.Email(contextDO, user.Email);

                                systemRepository_bankAccount.AddEntity<Model.Models.Email>(email);
                            }

                            if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                            {
                                Model.Models.EserviceNotification statusNotification = new Model.Models.EserviceNotification(paymentRequest);

                                systemRepository_bankAccount.AddEntity<Model.Models.EserviceNotification>(statusNotification);
                            }
                        }

                        unitOfWork_bankAccount.Save();
                    }
                }
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.ProcessTransactionFiles)
                    .Log(LogLevel.Error, "Job was canceled due to a token cancellation request;", ex);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.ProcessTransactionFiles)
                    .Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}", ex);
            }
        }

        public void Dispose()
        {
            if (!this.disposed)
            {
                this.disposed = true;

                JobLogger.Get(JobName.ProcessTransactionFiles).Log(LogLevel.Info, "ProcessTransactionFiles job disposed");
            }
        }
    }

    public class PendingPaymentRequest
    {
        public PaymentRequest PaymentRequest { get; set; }
        public bool IsReferenced { get; set; }

        public PendingPaymentRequest(PaymentRequest paymentRequest, bool isReferenced)
        {
            this.PaymentRequest = paymentRequest;
            this.IsReferenced = isReferenced;
        }
    }
}
