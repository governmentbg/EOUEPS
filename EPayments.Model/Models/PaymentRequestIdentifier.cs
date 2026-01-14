using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class PaymentRequestIdentifier
    {
        public int PaymentRequestIdentifierId { get; set; }
        public DateTime Date { get; set; }
        public int Counter { get; set; }
    }

    public class PaymentRequestIdentifierMap : EntityTypeConfiguration<PaymentRequestIdentifier>
    {
        public PaymentRequestIdentifierMap()
        {
            // Primary Key
            this.HasKey(t => t.PaymentRequestIdentifierId);

            this.Property(t => t.PaymentRequestIdentifierId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Table & Column Mappings
            this.ToTable("PaymentRequestIdentifiers");
            this.Property(t => t.PaymentRequestIdentifierId).HasColumnName("PaymentRequestIdentifierId");
            this.Property(t => t.Date).HasColumnName("Date");
            this.Property(t => t.Counter).HasColumnName("Counter");
        }
    }
}
