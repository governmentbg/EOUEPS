using EPayments.Model.Models;
using System;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class DistribtutionTypeVO
    {
        public int DistributionTypeId { get; set; }

        public string Name { get; set; }

        public static Expression<Func<DistributionType, DistribtutionTypeVO>> Map { get; } = (dt) =>
           new DistribtutionTypeVO()
           {
               DistributionTypeId = dt.DistributionTypeId,
               Name = dt.Name
           };
    }
}