using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class DistributionType
    {
        public int DistributionTypeId { get; set; }
        
        public string Name { get; set; }
        
        public virtual ICollection<EserviceClient> Clients { get; set; } = new HashSet<EserviceClient>();
        
        public virtual ICollection<DistributionRevenue> DistributionRevenues { get; set; } = new HashSet<DistributionRevenue>();
    }

    public class DistributionTypeMap : EntityTypeConfiguration<DistributionType>
    {
        public DistributionTypeMap()
        {
            this.HasKey(t => t.DistributionTypeId);
            
            this.Property(t => t.DistributionTypeId).HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);
            
            this.Property(t => t.DistributionTypeId).HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(50);
            
            this.HasMany(t => t.DistributionRevenues)
                .WithRequired(dr => dr.DistributionType)
                .HasForeignKey(dr => dr.DistributionTypeId);
            
            this.ToTable("DistributionTypes");
            this.Property(t => t.DistributionTypeId).HasColumnName("DistributionTypeId");
            this.Property(t => t.Name).HasColumnName("Name");
        }
    }
}
