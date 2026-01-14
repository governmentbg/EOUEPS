using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Models;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Shared;
using System.Collections.Generic;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class PaymentRequestsVM
    {
        public IList<ProcessedRequestVO> Requests { get; set; }
        public PagingVM RequestsPagingOptions { get; set; }
        public EserviceAdminRequestSearchDO SearchDO { get; set; }
        public string EserviceClientName { get; set; } 
        public string EserviceObligationType { get; set; }
        public string EserviceObligationTypeCode { get; set; }
        public List<PaymentRequestStatusModel> PaymentRequestStatuses { get; set; }
    }
}