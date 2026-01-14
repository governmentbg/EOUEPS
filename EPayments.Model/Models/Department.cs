using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class Department
    {
        public int DepartmentId { get; set; }
        
        public string Name { get; set; }
        
        public string UniqueIdentificationNumber { get; set; }
        
        public string UnifiedBudgetClassifier { get; set; }

        public bool IsActive { get; set; }

        public virtual ICollection<EserviceClient> EserviceClients { get; set; } = new HashSet<EserviceClient>(); 
    }

    public class DepartmentMap : EntityTypeConfiguration<Department>
    {
        public DepartmentMap()
        {
            // Primary Key
            this.HasKey(t => t.DepartmentId);

            this.Property(t => t.DepartmentId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(200);

            this.Property(t => t.UniqueIdentificationNumber)
                .HasMaxLength(16);

            this.Property(t => t.UnifiedBudgetClassifier)
                .HasMaxLength(16);

            // Table & Column Mappings
            this.ToTable("Departments");
            this.Property(t => t.DepartmentId).HasColumnName("DepartmentId");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.UniqueIdentificationNumber).HasColumnName("UniqueIdentificationNumber");
            this.Property(t => t.UnifiedBudgetClassifier).HasColumnName("UnifiedBudgetClassifier");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
        }
    }
}
