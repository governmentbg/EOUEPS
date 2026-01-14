using EPayments.Model.Enums;
using System;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class InternalAdminUser
    {
        public int InternalAdminUserId { get; set; }
        public string Name { get; set; }
        public string Egn { get; set; }
        public bool IsSuperadmin { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreateDate { get; set; }
        public InternalAdminUserPermissionEnum? Permissions { get; set; }
    }

    public class InternalAdminUserMap : EntityTypeConfiguration<InternalAdminUser>
    {
        public InternalAdminUserMap()
        {
            // Primary Key
            this.HasKey(t => t.InternalAdminUserId);

            // Properties
            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.Egn)
                .IsRequired()
                .HasMaxLength(10);

            // Table & Column Mappings
            this.ToTable("InternalAdminUsers");
            this.Property(t => t.InternalAdminUserId).HasColumnName("InternalAdminUserId");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.Egn).HasColumnName("Egn");
            this.Property(t => t.IsSuperadmin).HasColumnName("IsSuperadmin");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
            this.Property(t => t.Permissions).HasColumnName("Permissions");
        }
    }
}
