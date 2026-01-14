using EPayments.Data.Repositories.Interfaces;
using System.Web.Http;
using System.Web.Http.Controllers;
using System.Net.Http;
using Autofac.Integration.WebApi;
using Autofac;
using System.Web;
using EPayments.Common.Helpers;

namespace EPayments.Service.Api.Auth.Attributes
{
    public abstract class HmacAuthorizeAttribute : AuthorizeAttribute
    {
        protected override bool IsAuthorized(HttpActionContext actionContext)
        {
            try
            {
                var dependencyScope = actionContext.Request.GetDependencyScope();
                var lifetimeScope = dependencyScope.GetRequestLifetimeScope();
                var systemRepository = lifetimeScope.Resolve<ISystemRepository>();

                bool isAuthorized = false;

                string postData = actionContext.Request.Content.ReadAsStringAsync().Result;

                string clientId = HttpUtility.ParseQueryString(postData)["clientId"];
                string hmac = HttpUtility.ParseQueryString(postData)["hmac"];
                string data = HttpUtility.ParseQueryString(postData)["data"];

                HmacClientDO clientDO = this.GetHmacClientDO(systemRepository, clientId);

                if (clientDO != null && clientDO.IsActive)
                {
                    string verificationHmac = HmacRequestHelper.CalculateHmac(clientDO.SecretKey, data);

                    isAuthorized = verificationHmac == hmac;
                }

                return isAuthorized;
            }
            catch
            {
                return false;
            }
        }

        protected abstract HmacClientDO GetHmacClientDO(ISystemRepository systemRepository, string clientId);
    }
}
