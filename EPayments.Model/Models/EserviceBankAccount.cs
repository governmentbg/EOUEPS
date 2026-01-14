using System.Data.Entity.ModelConfiguration;

namespace EPayments.Model.Models
{
    public partial class EserviceBankAccount
    {
        public int EserviceBankAccountId { get; set; }
        public int BankId { get; set; }
        public string Iban { get; set; }
        public bool UploadTransactions { get; set; }
        public string TransactionsFilesPathUnread { get; set; }
        public string TransactionsFilesPathRead { get; set; }
        public bool IsActive { get; set; }

        public virtual Bank Bank { get; set; }
    }

    public class EserviceBankAccountMap : EntityTypeConfiguration<EserviceBankAccount>
    {
        public EserviceBankAccountMap()
        {
            // Primary Key
            this.HasKey(t => t.EserviceBankAccountId);

            // Properties
            this.Property(t => t.Iban)
                .IsRequired()
                .HasMaxLength(50);

            // Table & Column Mappings
            this.ToTable("EserviceBankAccounts");
            this.Property(t => t.EserviceBankAccountId).HasColumnName("EserviceBankAccountId");
            this.Property(t => t.BankId).HasColumnName("BankId");
            this.Property(t => t.Iban).HasColumnName("Iban");
            this.Property(t => t.UploadTransactions).HasColumnName("UploadTransactions");
            this.Property(t => t.TransactionsFilesPathUnread).HasColumnName("TransactionsFilesPathUnread");
            this.Property(t => t.TransactionsFilesPathRead).HasColumnName("TransactionsFilesPathRead");
            this.Property(t => t.IsActive).HasColumnName("IsActive");
        }
    }
}
