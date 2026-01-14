using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Models;
using System;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace EPayments.Web.Auth
{
    public class EserviceAuthorizeAttribute : AuthorizeAttribute
    {
        public ISystemRepository systemRepository  { get; set; }

        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            try
            {
                bool isAuthorized = false;

                string clientId = httpContext.Request["clientId"];
                string hmac = httpContext.Request["hmac"];
                string data = httpContext.Request["data"];

                if (!String.IsNullOrWhiteSpace(clientId) && !String.IsNullOrWhiteSpace(hmac) && !String.IsNullOrWhiteSpace(data))
                {
                    EserviceClient client = systemRepository.GetEserviceClientByClientId(clientId);

                    if (client != null && client.IsActive)
                    {
                        string verificationHmac = HmacRequestHelper.CalculateHmac(client.SecretKey, data);

                        isAuthorized = verificationHmac == hmac;
                    }
                }

                return isAuthorized;
            }
            catch
            {
                return false;
            }
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