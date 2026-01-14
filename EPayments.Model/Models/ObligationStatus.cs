using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class ObligationStatus
    {
        public int ObligationStatusId { get; set; }

        public string Alias { get; set; }

        public string StatusText { get; set; }

        public virtual ICollection<PaymentRequestObligationLog> Obligations { get; set; } = new HashSet<PaymentRequestObligationLog>();
    }

    public class ObligationStatusMap : EntityTypeConfiguration<ObligationStatus>
    {
        public ObligationStatusMap()
        {
            this.HasKey(t => t.ObligationStatusId);

            this.Property(t => t.ObligationStatusId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            this.Property(t => t.Alias)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.StatusText)
                .IsRequired()
                .HasMaxLength(50);

            this.ToTable("ObligationStatuses");
            this.Property(t => t.ObligationStatusId).HasColumnName("ObligationStatusId");
            this.Property(t => t.Alias).HasColumnName("Alias");
            this.Property(t => t.StatusText).HasColumnName("StatusText");
        }
    }
}
