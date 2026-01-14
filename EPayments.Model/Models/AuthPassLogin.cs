using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class AuthPassLogin
    {
        public int AuthPassLoginId { get; set; }
        public Guid Gid { get; set; }
        public string IP { get; set; }
        public string PostData { get; set; }
        public DateTime LogDate { get; set; }
    }

    public class AuthPassLoginMap : EntityTypeConfiguration<AuthPassLogin>
    {
        public AuthPassLoginMap()
        {
            // Primary Key
            this.HasKey(t => t.AuthPassLoginId);

            this.Property(t => t.AuthPassLoginId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.IP)
                .HasMaxLength(50);

            this.Property(t => t.PostData)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("AuthPassLogins");
            this.Property(t => t.AuthPassLoginId).HasColumnName("AuthPassLoginId");
            this.Property(t => t.Gid).HasColumnName("Gid");
            this.Property(t => t.IP).HasColumnName("IP");
            this.Property(t => t.PostData).HasColumnName("PostData");
            this.Property(t => t.LogDate).HasColumnName("LogDate");
        }
    }
}
