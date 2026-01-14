using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class PaymentRequestStatusModel
    {
        public int PaymentRequestStatusId { get; set; }
        public string Alias { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public class PaymentRequestStatusMap : EntityTypeConfiguration<PaymentRequestStatusModel>
    {
        public PaymentRequestStatusMap()
        {
            // Primary Key
            this.HasKey(t => t.PaymentRequestStatusId);

            this.Property(t => t.PaymentRequestStatusId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.Alias)
                .IsRequired()
                .HasMaxLength(50);

            // Table & Column Mappings
            this.ToTable("PaymentRequestStatuses");
            this.Property(t => t.PaymentRequestStatusId).HasColumnName("PaymentRequestStatusId");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.Alias).HasColumnName("Alias");
            this.Property(t => t.Description).HasColumnName("Description");
        }
    }
}
