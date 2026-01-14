using EPayments.Data.ViewObjects.Web;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class TransactionListVM
    {
        public Dictionary<string, List<string>> BankAndIbanDictionary { get; set; }

        public List<SelectListItem> BankAccounts { get; set; }

        public IList<TransactionRecordVO> TransactionRecords { get; set; }
        public PagingVM TransactionRecordsPagingOptions { get; set; }
        public TransactionListSearchDO SearchDO { get; set; }

        public decimal TotalAmount { get; set; }
    }
}