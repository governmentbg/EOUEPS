using EPayments.Model.DataObjects.EmailTemplateContext;
using EPayments.Model.Enums;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class Email
    {
        public int EmailId { get; set; }
        public string Recipient { get; set; }
        public string MailTemplateName { get; set; }
        public string Context { get; set; }
        public NotificationStatus NotificationStatusId { get; set; }
        public int FailedAttempts { get; set; }
        public string FailedAttemptsErrors { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public byte[] Version { get; set; }

        private Email()
        {
        }

        public Email(FeedbackContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.FeedbackMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(NewPaymentRequestContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.NewPaymentRequestMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(StatusChangedPaymentRequestContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.StatusChangedPaymentRequestMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(StatusChangedObligationContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.StatusChangedObligationMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(CertificateExpirationContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.CertificateExpirationMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(SharePaymentContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.SharePaymentMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(AccessCodeActivatedContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.AccessCodeActivatedMessage.ToString(), JObject.FromObject(contextDO));
        }

        public Email(AccessCodeApplicantActivatedContextDO contextDO, string recipient)
        {
            SetValues(recipient, EmailTemplate.AccessCodeApplicantActivatedMessage.ToString(), JObject.FromObject(contextDO));
        }

        private void SetValues(string recipient, string mailTemplateName, JObject context)
        {
            var currentDate = DateTime.Now;

            this.Recipient = recipient;
            this.MailTemplateName = mailTemplateName;
            if (context != null)
            {
                this.Context = context.ToString();
            }
            this.NotificationStatusId = NotificationStatus.Pending;
            this.FailedAttempts = 0;
            this.CreateDate = currentDate;
            this.ModifyDate = currentDate;
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
    }

    public class EmailMap : EntityTypeConfiguration<Email>
    {
        public EmailMap()
        {
            // Primary Key
            this.HasKey(t => t.EmailId);

            this.Property(t => t.EmailId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Recipient)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.MailTemplateName)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.Version)
                .IsRequired()
                .IsFixedLength()
                .HasMaxLength(8)
                .IsRowVersion();

            // Table & Column Mappings
            this.ToTable("Emails");
            this.Property(t => t.EmailId).HasColumnName("EmailId");
            this.Property(t => t.Recipient).HasColumnName("Recipient");
            this.Property(t => t.MailTemplateName).HasColumnName("MailTemplateName");
            this.Property(t => t.Context).HasColumnName("Context");
            this.Property(t => t.NotificationStatusId).HasColumnName("NotificationStatusId");
            this.Property(t => t.FailedAttempts).HasColumnName("FailedAttempts");
            this.Property(t => t.FailedAttemptsErrors).HasColumnName("FailedAttemptsErrors");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
            this.Property(t => t.ModifyDate).HasColumnName("ModifyDate");
            this.Property(t => t.Version).HasColumnName("Version");
        }
    }
}
