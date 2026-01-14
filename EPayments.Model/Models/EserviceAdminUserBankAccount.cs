using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EserviceAdminUserBankAccount
    {
        public int EserviceAdminUserBankAccountId { get; set; }
        public int EserviceAdminUserId { get; set; }
        public int EserviceBankAccountId { get; set; }
        public bool IsActive { get; set; }
    }

    public class EserviceAdminUserBankAccountMap : EntityTypeConfiguration<EserviceAdminUserBankAccount>
    {
        public EserviceAdminUserBankAccountMap()
        {
            // Primary Key
            this.HasKey(t => t.EserviceAdminUserBankAccountId);

            // Properties
            // Table & Column Mappings
            this.ToTable("EserviceAdminUserBankAccounts");
            this.Property(t => t.EserviceAdminUserBankAccountId).HasColumnName("EserviceAdminUserBankAccountId");
            this.Property(t => t.EserviceAdminUserId).HasColumnName("EserviceAdminUserId");
            this.Property(t => t.EserviceBankAccountId).HasColumnName("EserviceBankAccountId");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
        }
    }
}
