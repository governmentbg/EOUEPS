using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace EPayments.Data.ViewObjects.Admin
{
    public class DistributionRevenueVO
    {
        public int DistributionRevenueId { get; set; }

        public DateTime CreatedAt { get; set; }

        public DateTime? DistributedDate { get; set; }

        public bool IsDistributed { get; set; }

        public decimal TotalSum { get; set; }

        public int DistributionType { get; set; }

        public bool IsFileGenerated { get; set; }

        public string FileName { get; set; }

        public IEnumerable<string> Errors { get; set; }

        public static Expression<Func<DistributionRevenue, DistributionRevenueVO>> Map { get; } = (dr) =>
           new DistributionRevenueVO()
           {
               DistributionRevenueId = dr.DistributionRevenueId,
               CreatedAt = dr.CreatedAt,
               DistributedDate = dr.DistributedDate,
               IsDistributed = dr.IsDistributed,
               TotalSum = dr.TotalSum,
               DistributionType = dr.DistributionTypeId,
               IsFileGenerated = dr.IsFileGenerated,
               FileName = dr.FileName,
               Errors = dr.DistributionErrors.Select(de => de.Error)
           };
    }
}