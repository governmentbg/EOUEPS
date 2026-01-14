using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class PaymentRequestStatusLog
    {
        public int PaymentRequestStatusLogId { get; set; }
        public int PaymentRequestId { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }
        public DateTime ChangeDate { get; set; }
    }

    public class PaymentRequestStatusLogMap : EntityTypeConfiguration<PaymentRequestStatusLog>
    {
        public PaymentRequestStatusLogMap()
        {
            // Primary Key
            this.HasKey(t => t.PaymentRequestStatusLogId);

            this.Property(t => t.PaymentRequestStatusLogId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            // Table & Column Mappings
            this.ToTable("PaymentRequestStatusLogs");
            this.Property(t => t.PaymentRequestStatusLogId).HasColumnName("PaymentRequestStatusLogId");
            this.Property(t => t.PaymentRequestId).HasColumnName("PaymentRequestId");
            this.Property(t => t.PaymentRequestStatusId).HasColumnName("PaymentRequestStatusId");
            this.Property(t => t.ChangeDate).HasColumnName("ChangeDate");
        }
    }
}
