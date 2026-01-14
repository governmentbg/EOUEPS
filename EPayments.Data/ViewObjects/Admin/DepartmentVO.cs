using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class DepartmentVO
    {
        public int DepartmentId { get; set; }
        public string Name { get; set; }
        public string UniqueIdentificationNumber { get; set; }
        public string UnifiedBudgetClassifier { get; set; }

        public static Expression<Func<Department, DepartmentVO>> CreateMap { get; } = dep =>
           new DepartmentVO()
           {
               DepartmentId = dep.DepartmentId,
               Name = dep.Name,
               UniqueIdentificationNumber = dep.UniqueIdentificationNumber,
               UnifiedBudgetClassifier = dep.UnifiedBudgetClassifier
           };
    }
}
