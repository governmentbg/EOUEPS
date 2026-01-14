using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class BoricaTransactionStatus
    {
        public int BoricaTransactionStatusId { get; set; }

        public string StatusText { get; set; }

        public string Alias { get; set; }

        public ICollection<BoricaTransaction> BoricaTransactions { get; set; } = new HashSet<BoricaTransaction>();
    }

    public class BoricaTransactionStatusMap : EntityTypeConfiguration<BoricaTransactionStatus>
    {
        public BoricaTransactionStatusMap()
        {
            this.HasKey(t => t.BoricaTransactionStatusId);

            this.Property(t => t.BoricaTransactionStatusId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            this.Property(t => t.StatusText)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.Alias)
                .IsRequired()
                .HasMaxLength(100);

            this.ToTable("BoricaTransactionStatuses");
            this.Property(t => t.BoricaTransactionStatusId).HasColumnName("BoricaTransactionStatusId");
            this.Property(t => t.StatusText).HasColumnName("StatusText");
            this.Property(t => t.Alias).HasColumnName("Alias");
        }
    }
}
