using System;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security.Cookies;
using Owin;
using System.Web.Configuration;

namespace EPayments.Admin
{
    public partial class Startup
    {
        // For more information on configuring authentication, please visit http://go.microsoft.com/fwlink/?LinkId=301864
        public void ConfigureAuth(IAppBuilder app)
        {
            app.CreatePerOwinContext<EPaymentsAdmUserManager>(EPaymentsAdmUserManager.Create);
            app.CreatePerOwinContext<EPaymentsAdmSignInManager>(EPaymentsAdmSignInManager.Create);

            app.UseCookieAuthentication(new CookieAuthenticationOptions
            {
                AuthenticationType = DefaultAuthenticationTypes.ApplicationCookie,
                LoginPath = new PathString("/Account/Login"),
                //SlidingExpiration = false,
                ExpireTimeSpan = ((SessionStateSection)WebConfigurationManager.GetSection("system.web/sessionState")).Timeout,
                //LogoutPath = new PathString("/Account/Logout"),
                CookieName = "_epaymentsadm_identity_",
                Provider = new CookieAuthenticationProvider { }
            });
        }
    }
}