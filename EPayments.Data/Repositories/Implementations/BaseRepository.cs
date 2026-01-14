using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Implementations
{
    internal class BaseRepository : IBaseRepository
    {
        protected UnitOfWork unitOfWork;

        public BaseRepository(IUnitOfWork unitOfWork)
        {
            this.unitOfWork = (UnitOfWork)unitOfWork;
        }

        public void AddEntity<T>(T obj) where T : class
        {
            this.unitOfWork.DbContext.Set<T>().Add(obj);
        }
    }
}
