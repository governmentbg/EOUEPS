using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security;

namespace EPayments.Web
{
    public static class ClaimKeys
    {
        // common
        public const string Name = "user/name";
        public const string Uin = "user/uin";
        public const string UserId = "user/userId";
        public const string IsAuthorizedByAccessCode = "user/isAuthorizedByAccessCode";
        public const string AccessCode = "user/accessCode";
        public const string EserviceAdminId = "user/eserviceAdminId";
        public const string EserviceAdminDepartment = "user/eserviceAdminDepartment";
    }

    public class EPaymentsUser : IdentityUser
    {
        public string Name { get; set; }
        public string Uin { get; set; }
        public int? UserId { get; set; }
        public bool IsAuthorizedByAccessCode { get; set; }
        public string AccessCode { get; set; }
        public int? EserviceAdminId { get; set; }
        public string EserviceAdminDepartment { get; set; }

        public virtual Task<ClaimsIdentity> GenerateUserIdentityAsync(EPaymentsUserManager manager, string authenticationType)
        {
            var userIdentity = new ClaimsIdentity(authenticationType);

            userIdentity.AddClaim(new Claim(ClaimKeys.Name, this.Name ?? string.Empty));
            userIdentity.AddClaim(new Claim(ClaimKeys.Uin, this.Uin ?? string.Empty));
            userIdentity.AddClaim(new Claim(ClaimKeys.UserId, this.UserId.HasValue ? this.UserId.Value.ToString() : string.Empty));
            userIdentity.AddClaim(new Claim(ClaimKeys.IsAuthorizedByAccessCode, this.IsAuthorizedByAccessCode.ToString()));
            userIdentity.AddClaim(new Claim(ClaimKeys.AccessCode, this.AccessCode ?? string.Empty));
            userIdentity.AddClaim(new Claim(ClaimKeys.EserviceAdminId, this.EserviceAdminId.ToString()));
            userIdentity.AddClaim(new Claim(ClaimKeys.EserviceAdminDepartment, this.EserviceAdminDepartment ?? string.Empty));

            return Task.FromResult(userIdentity);
        }
    }

    public class EPaymentsUserManager : UserManager<EPaymentsUser>
    {
        public EPaymentsUserManager(IUserStore<EPaymentsUser> store) : base(store) { }

        public static EPaymentsUserManager Create(IdentityFactoryOptions<EPaymentsUserManager> options, IOwinContext context)
        {
            var manager = new EPaymentsUserManager(new UserStore<EPaymentsUser>());
            var dataProtectionProvider = options.DataProtectionProvider;
            if (dataProtectionProvider != null)
            {
                manager.UserTokenProvider = new DataProtectorTokenProvider<EPaymentsUser>(dataProtectionProvider.Create("_epayments_identity_"));
            }

            return manager;
        }

        public static EPaymentsUser LoadUser(ClaimsIdentity identity)
        {
            return new EPaymentsUser()
            {
                Name = identity.FindFirstValue(ClaimKeys.Name),
                Uin = identity.FindFirstValue(ClaimKeys.Uin),
                UserId = !String.IsNullOrWhiteSpace(identity.FindFirstValue(ClaimKeys.UserId)) ? int.Parse(identity.FindFirstValue(ClaimKeys.UserId)) : (int?)null,
                IsAuthorizedByAccessCode = bool.Parse(identity.FindFirstValue(ClaimKeys.IsAuthorizedByAccessCode)),
                AccessCode = identity.FindFirstValue(ClaimKeys.AccessCode),
                EserviceAdminId = !String.IsNullOrWhiteSpace(identity.FindFirstValue(ClaimKeys.EserviceAdminId)) ? int.Parse(identity.FindFirstValue(ClaimKeys.EserviceAdminId)) : (int?)null,
                EserviceAdminDepartment = identity.FindFirstValue(ClaimKeys.EserviceAdminDepartment),
            };
        }
    }

    public class EPaymentsSignInManager : SignInManager<EPaymentsUser, string>
    {
        public EPaymentsSignInManager(EPaymentsUserManager userManager, IAuthenticationManager authenticationManager)
            : base(userManager, authenticationManager)
        {
            this.AuthenticationType = DefaultAuthenticationTypes.ApplicationCookie;
        }

        public static EPaymentsSignInManager Create(IdentityFactoryOptions<EPaymentsSignInManager> options, IOwinContext context)
        {
            return new EPaymentsSignInManager(context.GetUserManager<EPaymentsUserManager>(), context.Authentication);
        }

        public override Task<ClaimsIdentity> CreateUserIdentityAsync(EPaymentsUser user)
        {
            return user.GenerateUserIdentityAsync((EPaymentsUserManager)UserManager, DefaultAuthenticationTypes.ApplicationCookie);
        }

        public override async Task SignInAsync(EPaymentsUser user, bool isPersistent, bool rememberBrowser)
        {
            var userIdentity = await CreateUserIdentityAsync(user);

            AuthenticationManager.SignOut();

            var authenticationProperties = new AuthenticationProperties() { IsPersistent = isPersistent };

            AuthenticationManager.SignIn(
                authenticationProperties,
                userIdentity);
        }
    }
}
