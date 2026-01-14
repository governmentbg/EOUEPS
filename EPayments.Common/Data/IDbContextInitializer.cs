using System.Data.Entity;

namespace EPayments.Common.Data
{
    public interface IDbContextInitializer
    {
        void InitializeContext(DbContext context);
    }
}
