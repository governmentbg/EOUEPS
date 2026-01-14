using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EserviceClient
    {
        public int EserviceClientId { get; set; }
        public Guid Gid { get; set; }
        public int? VposClientId { get; set; }
        public string Alias { get; set; }
        public string AisName { get; set; }
        public string ServiceName { get; set; }
        public string ServiceDescription { get; set; }
        public int DepartmentId { get; set; }
        public string AccountBank { get; set; }
        public string AccountBIC { get; set; }
        public string AccountIBAN { get; set; }
        public string ClientId { get; set; }
        public string SecretKey { get; set; }
        public string DskVposMerchantId { get; set; }
        public string DskVposMerchantPassword { get; set; }
        public string BoricaVposTerminalId { get; set; }
        public string BoricaVposMerchantId { get; set; }
        public string BoricaVposRequestSignCertFileName { get; set; }
        public string BoricaVposRequestSignCertPassword { get; set; }
        public DateTime? BoricaVposRequestSignCertValidTo { get; set; }
        public bool? BoricaVposRequestSignCertExpMailSend { get; set; }
        public bool? BoricaVposRequestSignCertExpHideAdminMsg { get; set; }
        public string FiBankVposAccessKeystoreFilename { get; set; }
        public string FiBankVposAccessKeystorePassword { get; set; }
        public bool IsEpayVposEnabled { get; set; }
        public bool IsAuthPassAuthorized { get; set; }
        public bool IsActive { get; set; }
        public int DistributionTypeId { get; set; }
        public DistributionType DistributionType { get; set; }
        public bool AggregateToParent { get; set; }
        public int? ParentId { get; set; }
        public virtual EserviceClient Parent { get; set; }
        public virtual VposClient VposClient { get; set; }
        public virtual Department Department { get; set; }
        //public int ObligationTypeId { get; set; }
       // public virtual ObligationType ObligationType { get; set; }
        public string DeliveryAdminstrationId { get; set; }
        public string DeliveryAdministrationGuid { get; set; }

        //public string BudgetCode { get; set; }

        public bool IsBoricaVposEnabled { get; set; } 
        public virtual ICollection<EserviceClient> Children { get; set; } = new HashSet<EserviceClient>();
        public virtual ICollection<DistributionRevenuePayment> DistributionRevenuePayments { get; set; } = new HashSet<DistributionRevenuePayment>();

        public ICollection<EserviceDeliveryNotification> EserviceDeliveryNotifications { get; set; } = new HashSet<EserviceDeliveryNotification>();
    }

    public class EserviceClientMap : EntityTypeConfiguration<EserviceClient>
    {
        public EserviceClientMap()
        {
            // Primary Key
            this.HasKey(t => t.EserviceClientId);

            this.Property(t => t.EserviceClientId)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);

            // Properties
            this.Property(t => t.ClientId)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.Alias)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.AisName)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.ServiceName)
                .HasMaxLength(200);

            this.Property(t => t.AccountBank)
                .IsRequired()
                .HasMaxLength(100);

            this.Property(t => t.AccountBIC)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.AccountIBAN)
                .IsRequired()
                .HasMaxLength(50);

            this.Property(t => t.SecretKey)
                .IsRequired()
                .HasMaxLength(200);

            this.Property(t => t.DskVposMerchantId)
                .HasMaxLength(100);

            //this.Property(t => t.BudgetCode)
            //    .HasMaxLength(10);

            //this.HasRequired(t => t.ObligationType)
            //    .WithMany(ot => ot.EserviceClients)
            //    .HasForeignKey(t => t.ObligationTypeId);

            this.HasRequired(t => t.DistributionType)
                .WithMany(dt => dt.Clients)
                .HasForeignKey(t => t.DistributionTypeId);

            this.HasOptional(t => t.Parent)
                .WithMany(d => d.Children)
                .HasForeignKey(t => t.ParentId);

            this.HasRequired(t => t.Department)
                .WithMany(d => d.EserviceClients)
                .HasForeignKey(t => t.DepartmentId);

            // Table & Column Mappings
            this.ToTable("EserviceClients");
            this.Property(t => t.EserviceClientId).HasColumnName("EserviceClientId");
            this.Property(t => t.Gid).HasColumnName("Gid");
            this.Property(t => t.VposClientId).HasColumnName("VposClientId");
            this.Property(t => t.Alias).HasColumnName("Alias");
            this.Property(t => t.AisName).HasColumnName("AisName");
            this.Property(t => t.ServiceName).HasColumnName("ServiceName");
            this.Property(t => t.ServiceDescription).HasColumnName("ServiceDescription");
            this.Property(t => t.DepartmentId).HasColumnName("DepartmentId");
            this.Property(t => t.AccountBank).HasColumnName("AccountBank");
            this.Property(t => t.AccountBIC).HasColumnName("AccountBIC");
            this.Property(t => t.AccountIBAN).HasColumnName("AccountIBAN");
            this.Property(t => t.ClientId).HasColumnName("ClientId");
            this.Property(t => t.SecretKey).HasColumnName("SecretKey");
            this.Property(t => t.DskVposMerchantId).HasColumnName("DskVposMerchantId");
            this.Property(t => t.DskVposMerchantPassword).HasColumnName("DskVposMerchantPassword");
            this.Property(t => t.BoricaVposRequestSignCertFileName).HasColumnName("BoricaVposRequestSignCertFileName");
            this.Property(t => t.BoricaVposRequestSignCertPassword).HasColumnName("BoricaVposRequestSignCertPassword");
            this.Property(t => t.BoricaVposRequestSignCertValidTo).HasColumnName("BoricaVposRequestSignCertValidTo");
            this.Property(t => t.BoricaVposRequestSignCertExpHideAdminMsg).HasColumnName("BoricaVposRequestSignCertExpHideAdminMsg");
            this.Property(t => t.BoricaVposRequestSignCertExpMailSend).HasColumnName("BoricaVposRequestSignCertExpMailSend");
            this.Property(t => t.FiBankVposAccessKeystoreFilename).HasColumnName("FiBankVposAccessKeystoreFilename");
            this.Property(t => t.FiBankVposAccessKeystorePassword).HasColumnName("FiBankVposAccessKeystorePassword");
            this.Property(t => t.IsEpayVposEnabled).HasColumnName("IsEpayVposEnabled");
            this.Property(t => t.IsAuthPassAuthorized).HasColumnName("IsAuthPassAuthorized");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
            //this.Property(t => t.ObligationTypeId).HasColumnName("ObligationTypeId");
            this.Property(t => t.AggregateToParent).HasColumnName("AggregateToParent");
            this.Property(t => t.DistributionTypeId).HasColumnName("DistributionTypeId");
            this.Property(t => t.DeliveryAdminstrationId).HasColumnName("DeliveryAdminstrationId");
            this.Property(t => t.DeliveryAdministrationGuid).HasColumnName("DeliveryAdministrationGuid");
            this.Property(t => t.IsBoricaVposEnabled).HasColumnName("IsBoricaVposEnabled");
            //this.Property(t => t.BudgetCode).HasColumnName("BudgetCode");
        }
    }
}
