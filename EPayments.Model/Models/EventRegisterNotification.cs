using EPayments.Common.Helpers;
using EPayments.Model.DataObjects.EmailTemplateContext;
using EPayments.Model.Enums;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;
using System.Text;

namespace EPayments.Model.Models
{
    public partial class EventRegisterNotification
    {
        public int EventRegisterNotificationId { get; set; }
        public int? PaymentRequestId { get; set; }
        public DateTime EventTime { get; set; }
        public string EventType { get; set; }
        public string EventDescription { get; set; }
        public string EventDocRegNumber { get; set; }
        public NotificationStatus NotificationStatusId { get; set; }
        public int FailedAttempts { get; set; }
        public string FailedAttemptsErrors { get; set; }
        public DateTime? SendNotBefore { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public byte[] Version { get; set; }

        private EventRegisterNotification()
        {
        }

        public EventRegisterNotification(int? paymentRequestId, EventRegisterNotificationType eventType, string eventDescription, string eventDocRegNumber)
        {
            this.PaymentRequestId = paymentRequestId;
            this.EventTime = DateTime.Now;
            this.EventType = eventType.ToString();
            this.EventDescription = eventDescription;
            this.EventDocRegNumber = eventDocRegNumber;
            this.NotificationStatusId = NotificationStatus.Pending;
            this.FailedAttempts = 0;
            this.CreateDate = this.ModifyDate = DateTime.Now;
            this.SendNotBefore = null;
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

    public class EventRegisterNotificationMap : EntityTypeConfiguration<EventRegisterNotification>
    {
        public EventRegisterNotificationMap()
        {
            // Primary Key
            this.HasKey(t => t.EventRegisterNotificationId);

            this.Property(t => t.EventRegisterNotificationId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Version)
                .IsRequired()
                .IsFixedLength()
                .HasMaxLength(8)
                .IsRowVersion();

            // Table & Column Mappings
            this.ToTable("EventRegisterNotifications");
            this.Property(t => t.EventRegisterNotificationId).HasColumnName("EventRegisterNotificationId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.EventTime).HasColumnName("EventTime");
            this.Property(t => t.EventType).HasColumnName("EventType");
            this.Property(t => t.EventDescription).HasColumnName("EventDescription");
            this.Property(t => t.EventDocRegNumber).HasColumnName("EventDocRegNumber");
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
