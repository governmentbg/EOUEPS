using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Models
{
    public class DistributionError
    {
        public int DistributionErrorId { get; set; }

        public string Error { get; set; }

        public DateTime CreatedAt { get; set; }

        public int DistributionRevenueId { get; set; }

        public virtual DistributionRevenue DistributionRevenue { get; set; }
    }

    public class DistributionErrorMap : EntityTypeConfiguration<DistributionError>
    {
        public DistributionErrorMap()
        {
            this.HasKey(t => t.DistributionErrorId);

            this.Property(t => t.DistributionErrorId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            this.Property(t => t.Error)
                .IsRequired()
                .HasMaxLength(500);

            this.ToTable("DistributionErrors");
            this.Property(t => t.DistributionErrorId).HasColumnName("DistributionErrorId");
            this.Property(t => t.Error).HasColumnName("Error");
            this.Property(t => t.CreatedAt).HasColumnName("CreatedAt");
            this.Property(t => t.DistributionRevenueId).HasColumnName("DistributionRevenueId");
        }
    }
}