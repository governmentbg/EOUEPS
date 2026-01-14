using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class DistributionRevenue
    {
        public int DistributionRevenueId { get; set; }

        public DateTime CreatedAt { get; set; }

        public DateTime? DistributedDate { get; set; }

        public bool IsDistributed { get; set; }

        public decimal TotalSum { get; set; }

        public bool IsFileGenerated { get; set; }

        public string FileName { get; set; } 

        public int DistributionTypeId { get; set; }

        public virtual DistributionType DistributionType { get; set; }

        public virtual ICollection<DistributionRevenuePayment> DistributionRevenuePayments { get; set; } = new HashSet<DistributionRevenuePayment>();

        public virtual ICollection<DistributionError> DistributionErrors { get; set; } = new HashSet<DistributionError>();
    }

    public class DistributionRevenueMap : EntityTypeConfiguration<DistributionRevenue>
    {
        public DistributionRevenueMap()
        {
            this.HasKey(t => t.DistributionRevenueId);

            this.Property(t => t.DistributionRevenueId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            this.Property(t => t.FileName).HasMaxLength(50);

            this.HasMany(t => t.DistributionErrors)
                .WithRequired(de => de.DistributionRevenue)
                .HasForeignKey(de => de.DistributionRevenueId);

            this.ToTable("DistributionRevenues");
            this.Property(t => t.DistributionRevenueId).HasColumnName("DistributionRevenueId");
            this.Property(t => t.CreatedAt).HasColumnName("CreatedAt");
            this.Property(t => t.DistributedDate).HasColumnName("DistributedDate");
            this.Property(t => t.IsDistributed).HasColumnName("IsDistributed");
            this.Property(t => t.TotalSum).HasColumnName("TotalSum");
            this.Property(t => t.IsFileGenerated).HasColumnName("IsFileGenerated");
            this.Property(t => t.FileName).HasColumnName("FileName");
            this.Property(t => t.DistributionTypeId).HasColumnName("DistributionTypeId");
        }
    }
}
