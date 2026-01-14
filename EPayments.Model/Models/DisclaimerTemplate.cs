using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class DisclaimerTemplate
    {
        public int DisclaimerTemplateId { get; set; }
        public string DisclaimerText { get; set; }
        public DateTime CreateDate { get; set; }
        public bool IsActive { get; set; }
    }

    public class DisclaimerTemplateMap : EntityTypeConfiguration<DisclaimerTemplate>
    {
        public DisclaimerTemplateMap()
        {
            // Primary Key
            this.HasKey(t => t.DisclaimerTemplateId);

            this.Property(t => t.DisclaimerTemplateId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.DisclaimerText)
                .IsRequired();

            // Table & Column Mappings
            this.ToTable("DisclaimerTemplates");
            this.Property(t => t.DisclaimerTemplateId).HasColumnName("DisclaimerTemplateId");
            this.Property(t => t.DisclaimerText).HasColumnName("DisclaimerText");
            this.Property(t => t.CreateDate).HasColumnName("CreateDate");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
        }
    }
}
