using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class ErrorLog
    {
        public int ErrorLogId { get; set; }
        public DateTime LogDate { get; set; }
        public string IP { get; set; }
        public string ControllerName { get; set; }
        public string ActionName { get; set; }
        public Guid? RequestId { get; set; }
        public string RequestInfo { get; set; }
        public string ErrorInfo { get; set; }
    }

    public class ErrorLogMap : EntityTypeConfiguration<ErrorLog>
    {
        public ErrorLogMap()
        {
            // Primary Key
            this.HasKey(t => t.ErrorLogId);

            this.Property(t => t.ErrorLogId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.IP)
                .HasMaxLength(50);

            this.Property(t => t.ControllerName)
                .HasMaxLength(50);

            this.Property(t => t.ActionName)
                .HasMaxLength(100);

            // Table & Column Mappings
            this.ToTable("ErrorLogs");
            this.Property(t => t.ErrorLogId).HasColumnName("ErrorLogId");
            this.Property(t => t.LogDate).HasColumnName("LogDate");
            this.Property(t => t.IP).HasColumnName("IP");
            this.Property(t => t.ControllerName).HasColumnName("ControllerName");
            this.Property(t => t.ActionName).HasColumnName("ActionName");
            this.Property(t => t.RequestId).HasColumnName("RequestId");
            this.Property(t => t.RequestInfo).HasColumnName("RequestInfo");
            this.Property(t => t.ErrorInfo).HasColumnName("ErrorInfo");
        }
    }
}
