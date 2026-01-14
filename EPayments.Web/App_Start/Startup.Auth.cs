using System;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security.Cookies;
using Owin;
using System.Web.Configuration;

namespace EPayments.Web
{
    public partial class Startup
    {
        // For more information on configuring authentication, please visit http://go.microsoft.com/fwlink/?LinkId=301864
        public void ConfigureAuth(IAppBuilder app)
        {
            app.CreatePerOwinContext<EPaymentsUserManager>(EPaymentsUserManager.Create);
            app.CreatePerOwinContext<EPaymentsSignInManager>(EPaymentsSignInManager.Create);

            app.UseCookieAuthentication(new CookieAuthenticationOptions
            {
                AuthenticationType = DefaultAuthenticationTypes.ApplicationCookie,
                //LoginPath = new PathString("/Account/Login"),
                LoginPath = new PathString("/Home/Index"),
                //SlidingExpiration = false,
                ExpireTimeSpan = ((SessionStateSection)WebConfigurationManager.GetSection("system.web/sessionState")).Timeout,
                //LogoutPath = new PathString("/Account/Logout"),
                CookieName = "_epayments_identity_",
                Provider = new CookieAuthenticationProvider { }
            });
        }
    }
}