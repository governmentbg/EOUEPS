using EPayments.Data.ViewObjects.Web;
using System.Collections.Generic;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class TransactionListPrintHtmlVM
    {
        public bool PrintResult { get; set; }

        public IList<TransactionRecordVO> Records{ get; set; }
    }
}