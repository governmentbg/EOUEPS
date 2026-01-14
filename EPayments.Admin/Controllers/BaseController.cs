using System.Security.Claims;
using System.Web.Mvc;

namespace EPayments.Admin.Controllers
{
    public abstract partial class BaseController : Controller
    {
        public BaseController()
        {
        }

        private EPaymentsAdmUser _currentUser;
        protected EPaymentsAdmUser CurrentUser
        {
            get
            {
                if (_currentUser == null)
                {
                    ClaimsIdentity ci = this.User.Identity as ClaimsIdentity;
                    _currentUser = EPaymentsAdmUserManager.LoadUser(ci);
                }
                return _currentUser;
            }
        }
    }
}