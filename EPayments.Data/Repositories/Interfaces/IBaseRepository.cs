using EPayments.Data.ViewObjects;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IBaseRepository
    {
        void AddEntity<T>(T obj) where T : class;
    }
}
