using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class GlobalValue
    {
        public int GlobalValueId { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
        public DateTime ModifyDate { get; set; }
    }

    public class GlobalValueMap : EntityTypeConfiguration<GlobalValue>
    {
        public GlobalValueMap()
        {
            // Primary Key
            this.HasKey(t => t.GlobalValueId);

            this.Property(t => t.GlobalValueId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.Key)
                .IsRequired()
                .HasMaxLength(200);

            // Table & Column Mappings
            this.ToTable("GlobalValues");
            this.Property(t => t.GlobalValueId).HasColumnName("GlobalValueId");
            this.Property(t => t.Key).HasColumnName("Key");
            this.Property(t => t.Value).HasColumnName("Value");
            this.Property(t => t.ModifyDate).HasColumnName("ModifyDate");
        }
    }

    public enum GlobalValueKey
    {
        DatabaseVersion,
        EmailJobLastInvocationTime,
        EserviceNotificationJobLastInvocationTime,
        EventRegisterNotificationJobLastInvocationTime,
        ExpiredRequestJobLastInvocationTime,
        BoricaVposResponseSignCertFileName,
        BoricaVposResponseSignCertValidTo,
        BoricaVposResponseSignCertExpMailSend,
        BoricaDistributionJobLastInvocationTime,
        BoricaUnprocessedRequestsJobLastInvocationTime,
        CVPosTransactionJobInvocationTime,
        CVPosTransactionFixJobInvocationTime
    }
}
