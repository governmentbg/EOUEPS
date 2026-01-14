using EPayments.Common.Helpers;
using EPayments.Model.DataObjects;
using EPayments.Model.Enums;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;
using System.Linq;

namespace EPayments.Model.Models
{
    public partial class EserviceNotification
    {
        public int EserviceNotificationId { get; set; }
        public int PaymentRequestId { get; set; }
        public int EserviceClientId { get; set; }
        public string Url { get; set; }
        public string PostData { get; set; }
        public NotificationStatus NotificationStatusId { get; set; }
        public int FailedAttempts { get; set; }
        public string FailedAttemptsErrors { get; set; }
        public DateTime? SendNotBefore { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public byte[] Version { get; set; }

        public virtual EserviceClient EserviceClient { get; set; }

        public virtual PaymentRequest PaymentRequest { get; set; }

        private EserviceNotification()
        {
        }

        public EserviceNotification(PaymentRequest request, int? sendingDelayInMinutes = null)
        {
            bool isMateus = request.PayOrder.HasValue;
            bool isAlgorithm2 = false;
            if (request.ObligationType != null && request.ObligationType.AlgorithmId == 2)
            {
                isAlgorithm2 = true;
            }

            var option = new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver(), Formatting = Formatting.None };
            string jsonData = "";

            if (isAlgorithm2)
            {
                var dataParam = new NotificationPostDataJsonAlgorithm2()
                {
                    Id = request.PaymentRequestIdentifier,
                    Status = request.PaymentRequestStatusId.ToString(),
                    ChangeTime = Formatter.DateTimeToIso8601Format(request.PaymentRequestStatusChangeTime),
                    AisPaymentId = request.AisPaymentId,
                    PaymentReferenceNumber = request.PaymentReferenceNumber,
                    PaymentAmount = request.PaymentAmount.ToString("G"),
                };
                jsonData = JsonConvert.SerializeObject(dataParam, option);
            }
            else
            {
                var dataParam = new NotificationPostDataJson()
                {
                    Id = request.PaymentRequestIdentifier,
                    Status = request.PaymentRequestStatusId.ToString(),
                    ChangeTime = Formatter.DateTimeToIso8601Format(request.PaymentRequestStatusChangeTime)
                };
                jsonData = JsonConvert.SerializeObject(dataParam, option);
            }

            string jsonDataMateus = string.Empty;
            if (isMateus)
            {
                var paidPaymentRequestByTransation = new List<PaymentRequest>();
                if (request.BoricaTransactions.Any(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid))
                {
                    paidPaymentRequestByTransation = request.BoricaTransactions.FirstOrDefault(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid)?.PaymentRequests.Where(pr => pr.EserviceClientId == request.EserviceClientId).ToList();
                }
                
                var notificationDataList = new List<NotificationPostDataJsonMateus>();

                foreach (var payment in paidPaymentRequestByTransation)
                {
                    var MDTJson = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(payment.AdditionalInformation));
                    var dataParamMateus = new NotificationPostDataJsonMateus()
                    {
                        PaymentRequestIdentifier = payment.PaymentRequestIdentifier,
                        Status = payment.PaymentRequestStatusId.ToString(),
                        ChangeTime = Formatter.DateTimeToIso8601Format(payment.PaymentRequestStatusChangeTime),
                        TransactionIdentifier = payment.BoricaTransactions.FirstOrDefault(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.Paid)?.Order,
                        PaidInstalmentSum = MDTJson?.PaidInstalmentSum.ToString(),
                        DebtInstalmentId = MDTJson?.DebtInstalmentId,
                        PaidInterestSum = MDTJson?.PaidInterestSum.ToString(),
                        ApplicantUin = payment.ApplicantUin,
                        ТaxSubjectId = MDTJson?.taxSubjectId
                    };
                    notificationDataList.Add(dataParamMateus);
                }

                jsonDataMateus = JsonConvert.SerializeObject(notificationDataList, option);
            }

            this.PaymentRequestId = request.PaymentRequestId;
            this.EserviceClientId = request.EserviceClientId;
            this.Url = request.AdministrativeServiceNotificationURL;
            this.PostData = isMateus ? jsonDataMateus : jsonData;
            this.NotificationStatusId = NotificationStatus.Pending;
            this.FailedAttempts = 0;
            this.CreateDate = this.ModifyDate = DateTime.Now;
            this.SendNotBefore = sendingDelayInMinutes.HasValue ? DateTime.Now.AddMinutes(sendingDelayInMinutes.Value) : (DateTime?)null;
        }

        public void SetStatus(NotificationStatus notificationStatusId)
        {
            this.NotificationStatusId = notificationStatusId;
            this.ModifyDate = DateTime.Now;
        }

        public void IncrementFailedAttempts(string exception)
        {
            JObject fae;
            if (String.IsNullOrEmpty(this.FailedAttemptsErrors))
            {
                fae = new JObject();
            }
            else
            {
                fae = JObject.Parse(this.FailedAttemptsErrors);
            }
            fae.Add(this.FailedAttempts.ToString(), exception);
            this.FailedAttemptsErrors = fae.ToString();
            this.FailedAttempts++;
            this.ModifyDate = DateTime.Now;
        }

        public void SetNextSendingAttemptTime(DateTime? sendNotBefore, bool setStatusTerminated)
        {
            this.SendNotBefore = sendNotBefore;

            if (setStatusTerminated)
            {
                this.NotificationStatusId = NotificationStatus.Terminated;
            }
        }
    }

    public class EserviceNotificationMap : EntityTypeConfiguration<EserviceNotification>
    {
        public EserviceNotificationMap()
        {
            // Primary Key
            this.HasKey(t => t.EserviceNotificationId);

            this.Property(t => t.EserviceNotificationId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Url)
                .IsRequired();

            this.Property(t => t.PostData)
                .IsRequired();

            this.Property(t => t.Version)
                .IsRequired()
                .IsFixedLength()
                .HasMaxLength(8)
                .IsRowVersion();

            // Table & Column Mappings
            this.ToTable("EserviceNotifications");
            this.Property(t => t.EserviceNotificationId).HasColumnName("EserviceNotificationId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.EserviceClientId).HasColumnName("EserviceClientId");
            this.Property(t => t.Url).HasColumnName("Url");
            this.Property(t => t.PostData).HasColumnName("PostData");
            this.Property(t => t.NotificationStatusId).HasColumnName("NotificationStatusId");
            this.Property(t => t.FailedAttempts).HasColumnName("FailedAttempts");
            this.Property(t => t.FailedAttemptsErrors).HasColumnName("FailedAttemptsErrors");
            this.Property(t => t.SendNotBefore).HasColumnName("SendNotBefore");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
            this.Property(t => t.ModifyDate).HasColumnName("ModifyDate");
            this.Property(t => t.Version).HasColumnName("Version");
        }
    }
}
