using EPayments.Common.Data;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace EPayments.Admin.Auth
{
    public class AdminAuthorizeAttribute : AuthorizeAttribute
    {
        public AdminAuthorizeAttribute(InternalAdminUserPermissionEnum requiredPermission)
        {
            this.RequiredPermission = requiredPermission;
        }

        public IUnitOfWork UnitOfWork { get; set; }

        private InternalAdminUserPermissionEnum RequiredPermission { get; }

        private bool HasChangedPermissions { get; set; } = false; 

        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            bool isAuthorized = base.AuthorizeCore(httpContext);

            if (isAuthorized)
            {
                ClaimsIdentity ci = httpContext.User.Identity as ClaimsIdentity;
                EPaymentsAdmUser currentUser = EPaymentsAdmUserManager.LoadUser(ci);

                InternalAdminUser user = UnitOfWork.DbContext.Set<InternalAdminUser>()
                    .SingleOrDefault(u => u.InternalAdminUserId == currentUser.UserId);

                if (user == null || user.IsActive != true)
                {
                    return false;
                }

                if (user.IsSuperadmin)
                {
                    return true;
                }

                if ((int?)user.Permissions != (int?)currentUser.Permission)
                {
                    this.HasChangedPermissions = true;
                    return false;
                }

                isAuthorized = user.Permissions != null ? 
                    (this.RequiredPermission == (user.Permissions & this.RequiredPermission)) : 
                    false;
            }

            return isAuthorized;
        }

        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            if (this.HasChangedPermissions == true)
            {
                filterContext.Result = new RedirectToRouteResult(
                new RouteValueDictionary(
                    new
                    {
                        controller = MVC.Account.Name,
                        action = MVC.Account.ActionNames.Logout
                    })
                );
                return;
            }
            
            filterContext.Result = new RedirectToRouteResult(
                new RouteValueDictionary(
                    new
                    {
                        controller = MVC.Home.Name,
                        action = MVC.Home.ActionNames.RedirectToError,
                        id = "104",
                        isIisError = true
                    })
                );
        }
    }
}