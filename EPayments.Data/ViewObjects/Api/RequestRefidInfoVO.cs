using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Api
{
    public class RequestRefidInfoVO
    {
        public int DistributionRevenueId { get; set; }
        public decimal TotalSum { get; set; }
        public DateTime CreatedAt { get; set; }
        public string Order { get; set; }
        public decimal Amount { get; set; }
        public DateTime TransactionDate { get; set; }
        public string Rrn { get; set; }
        public IEnumerable<RequestPaymentInfoParsedVO> PaymentInfo { get; set; } = new HashSet<RequestPaymentInfoParsedVO>();
        public string DepartmentName { get; set; }
        public string DepartmentUniqueIdentificationNumber { get; set; }
    }
}
