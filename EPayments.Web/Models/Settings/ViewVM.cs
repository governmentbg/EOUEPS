using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Models;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Settings
{
    public class ViewVM
    {
        public string Email { get; set; }

        public bool RequestNotifications { get; set; }

        public bool StatusNotifications { get; set; }

        public bool StatusObligationNotifications { get; set; }

        public bool AccessCodeNotifications { get; set; }

        public IList<RequestAccessVO> RequestAccessList { get; set; }

        public PagingVM RequestAccessListPagingOptions { get; set; }

        public RequestAccessListSearchDO SearchDO { get; set; }
    }
}