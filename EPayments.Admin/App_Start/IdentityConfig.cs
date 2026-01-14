using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using EPayments.Model.Enums;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security;

namespace EPayments.Admin
{
    public static class ClaimKeys
    {
        // common
        public const string Name = "user/name";
        public const string Uin = "user/uin";
        public const string UserId = "user/userId";
    }

    public class EPaymentsAdmUser : IdentityUser
    {
        public string Name { get; set; }
        public string Uin { get; set; }
        public int? UserId { get; set; }
        public InternalAdminUserPermissionEnum? Permission { get; set; }

        public virtual Task<ClaimsIdentity> GenerateUserIdentityAsync(EPaymentsAdmUserManager manager, string authenticationType)
        {
            var userIdentity = new ClaimsIdentity(authenticationType);

            userIdentity.AddClaim(new Claim(ClaimKeys.Name, this.Name ?? string.Empty));
            userIdentity.AddClaim(new Claim(ClaimKeys.Uin, this.Uin ?? string.Empty));
            userIdentity.AddClaim(new Claim(ClaimKeys.UserId, this.UserId.HasValue ? this.UserId.Value.ToString() : string.Empty));
            
            if (this.Permission != null)
            {
                userIdentity.AddClaim(new Claim(nameof(Permission), ((int)this.Permission).ToString()));
            }

            return Task.FromResult(userIdentity);
        }

        public bool HasPermission(InternalAdminUserPermissionEnum permission)
        {
            return this.Permission != null ? (this.Permission & permission) == permission : false;
        }
    }

    public class EPaymentsAdmUserManager : UserManager<EPaymentsAdmUser>
    {
        public EPaymentsAdmUserManager(IUserStore<EPaymentsAdmUser> store) : base(store) { }

        public static EPaymentsAdmUserManager Create(IdentityFactoryOptions<EPaymentsAdmUserManager> options, IOwinContext context)
        {
            var manager = new EPaymentsAdmUserManager(new UserStore<EPaymentsAdmUser>());
            var dataProtectionProvider = options.DataProtectionProvider;
            if (dataProtectionProvider != null)
            {
                manager.UserTokenProvider = new DataProtectorTokenProvider<EPaymentsAdmUser>(dataProtectionProvider.Create("_epaymentsadm_identity_"));
            }

            return manager;
        }

        public static EPaymentsAdmUser LoadUser(ClaimsIdentity identity)
        {
            string value = identity.FindFirstValue(nameof(EPaymentsAdmUser.Permission));
            InternalAdminUserPermissionEnum permission;

            return new EPaymentsAdmUser()
            {
                Name = identity.FindFirstValue(ClaimKeys.Name),
                Uin = identity.FindFirstValue(ClaimKeys.Uin),
                UserId = !String.IsNullOrWhiteSpace(identity.FindFirstValue(ClaimKeys.UserId)) ? int.Parse(identity.FindFirstValue(ClaimKeys.UserId)) : (int?)null, 
                Permission = (!string.IsNullOrWhiteSpace(value) && Enum.TryParse(value, out permission)) ? permission : (InternalAdminUserPermissionEnum?)null
            };
        }
    }

    public class EPaymentsAdmSignInManager : SignInManager<EPaymentsAdmUser, string>
    {
        public EPaymentsAdmSignInManager(EPaymentsAdmUserManager userManager, IAuthenticationManager authenticationManager)
            : base(userManager, authenticationManager)
        {
            this.AuthenticationType = DefaultAuthenticationTypes.ApplicationCookie;
        }

        public static EPaymentsAdmSignInManager Create(IdentityFactoryOptions<EPaymentsAdmSignInManager> options, IOwinContext context)
        {
            return new EPaymentsAdmSignInManager(context.GetUserManager<EPaymentsAdmUserManager>(), context.Authentication);
        }

        public override Task<ClaimsIdentity> CreateUserIdentityAsync(EPaymentsAdmUser user)
        {
            return user.GenerateUserIdentityAsync((EPaymentsAdmUserManager)UserManager, DefaultAuthenticationTypes.ApplicationCookie);
        }

        public override async Task SignInAsync(EPaymentsAdmUser user, bool isPersistent, bool rememberBrowser)
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
