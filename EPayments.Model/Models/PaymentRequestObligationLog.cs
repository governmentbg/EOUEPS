using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class PaymentRequestObligationLog
    {
        public int PaymentRequestObligationLogsId { get; set; }

        public int PaymentRequestId { get; set; }

        public virtual PaymentRequest PaymentRequest { get; set; }

        public int ObligationStatusId { get; set; }

        public virtual ObligationStatus ObligationStatus { get; set; }

        public DateTime ChangeDate { get; set; }

        public ICollection<EserviceDeliveryNotification> EserviceDeliveryNotifications { get; set; } = new HashSet<EserviceDeliveryNotification>(); 
    }

    public class PaymentRequestHistoryMap : EntityTypeConfiguration<PaymentRequestObligationLog>
    {
        public PaymentRequestHistoryMap()
        {
            this.HasKey(t => t.PaymentRequestObligationLogsId);

            this.Property(t => t.PaymentRequestObligationLogsId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            this.Property(t => t.ChangeDate)
                .IsRequired();

            this.HasRequired(t => t.PaymentRequest)
                .WithMany(pr => pr.PaymentRequestObligationLogs)
                .HasForeignKey(t => t.PaymentRequestId);

            this.HasRequired(t => t.ObligationStatus)
                .WithMany(phs => phs.Obligations)
                .HasForeignKey(t => t.ObligationStatusId);

            this.ToTable("PaymentRequestObligationLogs");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.ObligationStatusId).HasColumnName("ObligationStatusId");
            this.Property(t => t.ChangeDate).HasColumnName("ChangeDate");
        }
    }
}
