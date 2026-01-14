using EPayments.Model.Enums;
using Newtonsoft.Json.Linq;
using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EserviceDeliveryNotification
    {
        public int EserviceNotificationId { get; set; }
        public int PaymentRequestId { get; set; }
        public int PaymentRequestObligationLogsId { get; set; }
        public string Uniqueidentifier { get; set; }
        public string PersonUniqueIdentifier { get; set; }
        public int EserviceClientId { get; set; }
        public string ResponseCodes { get; set; }
        public NotificationStatus NotificationStatusId { get; set; }
        public int FailedAttempts { get; set; }
        public string FailedAttemptsErrors { get; set; }
        public DateTime? SendNotBefore { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public byte[] Version { get; set; }

        public virtual PaymentRequestObligationLog PaymentRequestObligationLog { get; set; }
        public virtual EserviceClient EserviceClient { get; set; }

        private EserviceDeliveryNotification()
        {
        }

        public EserviceDeliveryNotification(PaymentRequestObligationLog request, int? sendingDelayInMinutes = null) 
        {
            this.Uniqueidentifier = request.PaymentRequest.EserviceClient.DeliveryAdminstrationId;
            this.PersonUniqueIdentifier = request.PaymentRequest.ApplicantUin;
            this.PaymentRequestId = request.PaymentRequestId;
            this.EserviceClientId = request.PaymentRequest.EserviceClientId;
            this.PaymentRequestObligationLogsId = request.PaymentRequestObligationLogsId;
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

        public void AddResponseCodes(params string[] messages)
        {
            if (messages != null && messages.Length > 0)
            {
                this.ResponseCodes = string.Join("; ", messages);
            }
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

    public class EserviceDeliveryNotificationMap : EntityTypeConfiguration<EserviceDeliveryNotification>
    {
        public EserviceDeliveryNotificationMap()
        {
            // Primary Key
            this.HasKey(t => t.EserviceNotificationId);

            this.Property(t => t.EserviceNotificationId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Uniqueidentifier)
                .HasMaxLength(20);

            this.Property(t => t.PersonUniqueIdentifier)
                .HasMaxLength(20);

            this.Property(t => t.ResponseCodes)
                .HasMaxLength(100);

            this.Property(t => t.Version)
                .IsRequired()
                .IsFixedLength()
                .HasMaxLength(8)
                .IsRowVersion();

            this.HasRequired(t => t.PaymentRequestObligationLog)
                .WithMany(prol => prol.EserviceDeliveryNotifications)
                .HasForeignKey(t => t.PaymentRequestObligationLogsId);

            this.HasRequired(t => t.EserviceClient)
                .WithMany(ec => ec.EserviceDeliveryNotifications)
                .HasForeignKey(t => t.EserviceClientId);

            // Table & Column Mappings
            this.ToTable("EserviceDeliveryNotifications");
            this.Property(t => t.EserviceNotificationId).HasColumnName("EserviceNotificationId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.PaymentRequestObligationLogsId).HasColumnName("PaymentRequestObligationLogsId");
            this.Property(t => t.EserviceClientId).HasColumnName("EserviceClientId");
            this.Property(t => t.ResponseCodes).HasColumnName("ResponseCodes");
            this.Property(t => t.NotificationStatusId).HasColumnName("NotificationStatusId");
            this.Property(t => t.FailedAttempts).HasColumnName("FailedAttempts");
            this.Property(t => t.FailedAttemptsErrors).HasColumnName("FailedAttemptsErrors");
            this.Property(t => t.SendNotBefore).HasColumnName("SendNotBefore");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
            this.Property(t => t.ModifyDate).HasColumnName("ModifyDate");
            this.Property(t => t.Version).HasColumnName("Version");
            this.Property(t => t.Uniqueidentifier).HasColumnName("Uniqueidentifier");
            this.Property(t => t.PersonUniqueIdentifier).HasColumnName("PersonUniqueIdentifier");
        }
    }
}
