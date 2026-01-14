using System.Data.Entity;

namespace EPayments.Common.Data
{
    public interface IDbConfiguration
    {
        void AddConfiguration(DbModelBuilder modelBuilder);
    }
}
