using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public class DistributionRevenuePayment
    {
        [ForeignKey(nameof(PaymentRequest))]
        public int DistributionRevenuePaymentId { get; set; }

        public int EserviceClientId { get; set; }

        public virtual EserviceClient EserviceClient { get; set; }

        public virtual PaymentRequest PaymentRequest { get; set; } 

        public int DistributionRevenueId { get; set; } 

        public virtual DistributionRevenue DistributionRevenue { get; set; }

        public int BoricaTransactionId { get; set; }

        public virtual BoricaTransaction BoricaTransaction { get; set; } 

        public string DitribtutionError { get; set; }
    }

    public class DistributionRevenuePaymentsMap : EntityTypeConfiguration<DistributionRevenuePayment>
    {
        public DistributionRevenuePaymentsMap()
        {
            // Primary Key
            this.HasKey(t => t.DistributionRevenuePaymentId);

            this.HasRequired(t => t.EserviceClient)
                .WithMany(ec => ec.DistributionRevenuePayments)
                .HasForeignKey(t => t.EserviceClientId);

            this.HasRequired(t => t.DistributionRevenue)
                .WithMany(dr => dr.DistributionRevenuePayments)
                .HasForeignKey(t => t.DistributionRevenueId);

            this.HasRequired(t => t.PaymentRequest)
                .WithOptional(pr => pr.DistributionRevenuePayment);

            this.HasRequired(t => t.BoricaTransaction)
                .WithMany(bt => bt.DistributionRevenuePayments)
                .HasForeignKey(t => t.BoricaTransactionId);

            // Table & Column Mappings
            this.ToTable("DistributionRevenuePayments");
            this.Property(t => t.DistributionRevenuePaymentId).HasColumnName("DistributionRevenuePaymentId");
            this.Property(t => t.EserviceClientId).HasColumnName("EserviceClientId");
            this.Property(t => t.DistributionRevenueId).HasColumnName("DistributionRevenueId");
            this.Property(t => t.BoricaTransactionId).HasColumnName("BoricaTransactionId");
            this.Property(t => t.DitribtutionError).HasColumnName("DitribtutionError");
        }
    }
}
