using EPayments.Data.ViewObjects.Web;
using System.Collections.Generic;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class PrintPaymentRequestsVM
    {
        public bool PrintAllResults { get; set; }

        public IList<ProcessedRequestVO> Requests { get; set; }
    }
}