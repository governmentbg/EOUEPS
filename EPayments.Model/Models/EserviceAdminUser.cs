using System.Collections.Generic;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EserviceAdminUser
    {
        public int EserviceAdminUserId { get; set; }
        public string Username { get; set; }
        public string PasswordHash { get; set; }
        public string PasswordSalt { get; set; }
        public string IpAddressesForAccess { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public bool InsufficientAmountNotifications { get; set; }
        public bool OverpaidAmountNotifications { get; set; }
        public bool ReferencedNotifications { get; set; }
        public bool NotReferencedNotifications { get; set; }
        public bool IsActive { get; set; }
        public int? ReferringEserviceClientId { get; set; }
    }

    public class EserviceAdminUserMap : EntityTypeConfiguration<EserviceAdminUser>
    {
        public EserviceAdminUserMap()
        {
            // Primary Key
            this.HasKey(t => t.EserviceAdminUserId);

            // Properties
            this.Property(t => t.Username)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.PasswordHash)
                .IsRequired();

            this.Property(t => t.PasswordSalt)
                .IsRequired();

            this.Property(t => t.Name)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.Email)
                .HasMaxLength(100);

            // Table & Column Mappings
            this.ToTable("EserviceAdminUsers");
            this.Property(t => t.EserviceAdminUserId).HasColumnName("EserviceAdminUserId");
            this.Property(t => t.Username).HasColumnName("Username");
            this.Property(t => t.PasswordHash).HasColumnName("PasswordHash");
            this.Property(t => t.PasswordSalt).HasColumnName("PasswordSalt");
            this.Property(t => t.IpAddressesForAccess).HasColumnName("IpAddressesForAccess");
            this.Property(t => t.Name).HasColumnName("Name");
            this.Property(t => t.Email).HasColumnName("Email");
            this.Property(t => t.InsufficientAmountNotifications).HasColumnName("InsufficientAmountNotifications");
            this.Property(t => t.OverpaidAmountNotifications).HasColumnName("OverpaidAmountNotifications");
            this.Property(t => t.ReferencedNotifications).HasColumnName("ReferencedNotifications");
            this.Property(t => t.NotReferencedNotifications).HasColumnName("NotReferencedNotifications");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
            this.Property(t => t.ReferringEserviceClientId).HasColumnName("ReferringEserviceClientId");
        }
    }
}
