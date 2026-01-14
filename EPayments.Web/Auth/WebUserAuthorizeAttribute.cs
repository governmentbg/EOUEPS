using System.Security.Claims;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace EPayments.Web.Auth
{
    public class WebUserAuthorizeAttribute : AuthorizeAttribute
    {
        public bool AllowAuthorizationByAccessCode { get; set; }

        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            bool isAuthorized = base.AuthorizeCore(httpContext);

            if (isAuthorized)
            {
                ClaimsIdentity ci = httpContext.User.Identity as ClaimsIdentity;
                EPaymentsUser currentUser = EPaymentsUserManager.LoadUser(ci);

                if (currentUser.EserviceAdminId.HasValue)
                {
                    isAuthorized = false;
                }
                else if (!this.AllowAuthorizationByAccessCode)
                {
                    if (currentUser.IsAuthorizedByAccessCode)
                    {
                        isAuthorized = false;
                    }
                }
            }

            return isAuthorized;
        }

        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            filterContext.Result = new RedirectToRouteResult(
                new RouteValueDictionary(
                    new
                    {
                        controller = MVC.Home.Name,
                        action = MVC.Home.ActionNames.RedirectToError,
                        id = "401"
                    })
                );
        }
    }
}