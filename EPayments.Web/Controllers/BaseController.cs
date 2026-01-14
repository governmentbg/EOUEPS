using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;

namespace EPayments.Web.Controllers
{
    public abstract partial class BaseController : Controller
    {
        public BaseController()
        {
        }

        private EPaymentsUser _currentUser;
        protected EPaymentsUser CurrentUser
        {
            get
            {
                if (_currentUser == null)
                {
                    ClaimsIdentity ci = this.User.Identity as ClaimsIdentity;
                    _currentUser = EPaymentsUserManager.LoadUser(ci);
                }
                return _currentUser;
            }
        }
    }
}