using EPayments.EDelivery.Models;
using System;
using System.Runtime.Caching;

namespace EPayments.EDelivery.Manager
{
    public abstract class DeliveryRegisterBaseManager
    {
        private const string DepartmentsKey = "DepartmentsKey";
        private readonly MemoryCache memoryCache = MemoryCache.Default;
        private readonly CacheItemPolicy cacheItemPolicy = new CacheItemPolicy() { SlidingExpiration = TimeSpan.FromMinutes(5) };

        protected bool AddDepartmentsToCache(DepartmentInstitutionInfo[] departments)
        {
            return this.memoryCache.Add(DepartmentsKey, departments, this.cacheItemPolicy);
        }

        protected DepartmentInstitutionInfo[] GetDepartmentsFromCache()
        {
            try
            {
                return (DepartmentInstitutionInfo[])this.memoryCache.Get(DepartmentsKey);
            }
            catch (NotSupportedException)
            {
                return null;
            }
            catch (ArgumentNullException)
            {
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        protected bool ContainsDepartments()
        {
            try
            {
                return this.memoryCache.Contains(DepartmentsKey);
            }
            catch (NotSupportedException)
            {
                return false;
            }
            catch (ArgumentNullException)
            {
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
