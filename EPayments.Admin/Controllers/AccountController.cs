using EPayments.Admin.Common;
using EPayments.Admin.Models.Shared;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Common.Saml;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using System;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Xml;

namespace EPayments.Admin.Controllers
{
    public partial class AccountController : BaseController
    {
        private IAuthenticationManager AuthenticationManager
        {
            get
            {
                return HttpContext.GetOwinContext().Authentication;
            }
        }

        private EPaymentsAdmSignInManager SignInManager
        {
            get
            {
                return HttpContext.GetOwinContext().Get<EPaymentsAdmSignInManager>();
            }
        }

        private ISystemRepository systemRepository;
        private IUnitOfWork unitOfWork;

        public AccountController(ISystemRepository systemRepository, IUnitOfWork unitOfWork)
        {
            this.systemRepository = systemRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        [Authorize]
        public virtual ActionResult Logout()
        {
            AuthenticationManager.SignOut();

            TempData[TempDataKeys.ErrorMessage] = "За да прекратите сесията на електронния подпис и да излезете напълно от системата трябва да затворите всички прозорци на текущия браузер!";

            return RedirectToAction(MVC.Account.ActionNames.Login, MVC.Account.Name);
        }

        [HttpGet]
        public virtual ActionResult Login()
        {
            if (Request.IsAuthenticated)
            {
                return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
            }

            return View();
        }

        [HttpGet]
        public virtual ActionResult EAuth()
         {
            if (AppSettings.EPaymentsWeb_EAuthEnabled)
            {
                string returnUri = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Account.Name, MVC.Account.ActionNames.EAuthResponse).ToString();
                string metadataUri = AppSettings.EPaymentsWeb_EAuthMetadataUrl;

                XmlDocument authnRequest = SamlHelper.GenerateKEPAuthnRequest(
                    "Единен портал за електронни разплащания (admin system)",
                    AppSettings.EPaymentsWeb_EAuthProviderId,
                    returnUri,
                    AppSettings.EPaymentsWeb_EAuthRequestUrl,
                    AppSettings.EPaymentsWeb_EAuthExtServiceId,
                    AppSettings.EPaymentsWeb_EAuthExtProviderId,
                    metadataUri);
                var signedAuthnRequest = SamlHelper.SignXmlDocument(authnRequest, AppSettings.EPaymentsWeb_EAuthRequestSignCertificatePath, AppSettings.EPaymentsWeb_EAuthRequestSignCertificatePass);
                var base64AuthnRequest = Convert.ToBase64String(Encoding.UTF8.GetBytes(signedAuthnRequest));

                AutoPostVM model = new AutoPostVM();
                model.PostUrl = AppSettings.EPaymentsWeb_EAuthRequestUrl;
                model.PostValues = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>>();
                model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("SAMLRequest", base64AuthnRequest));
                model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("RelayState", String.Empty));

                return View(MVC.Shared.Views._AutoPost, model);
            }
            else 
            {
                InternalAdminUser internalAdminUser = this.systemRepository.GetInternalAdminUserByEgn("7409228750");
                EPaymentsAdmUser admUser = new EPaymentsAdmUser()
                {
                    UserId = internalAdminUser.InternalAdminUserId,
                    Uin = internalAdminUser.Egn,
                    Name = internalAdminUser.Name,
                    Permission = internalAdminUser.IsSuperadmin ?
                          InternalAdminUserPermissionEnum.Modify | InternalAdminUserPermissionEnum.ViewReferences | InternalAdminUserPermissionEnum.DistributionReferences :
                          internalAdminUser.Permissions
                };

                var task = SignInManager.SignInAsync(admUser, false, false);
                task.Wait();
                return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
            }
        }

        [HttpPost]
        [AllowAnonymous]
        public virtual ActionResult EAuthResponse(string samlResponse, string relayState)
        {
            if (!string.IsNullOrEmpty(samlResponse))
            {
                var decodedStr = Encoding.UTF8.GetString(Convert.FromBase64String(samlResponse));
                EAuthLoginDataDO response = SamlHelper.ParseEAuthResponse(decodedStr, AppSettings.EPaymentsWeb_EAuthResponseSignCertificateThumbprint, AppSettings.EPaymentsWeb_EAuthResponseSignCertificateValidateExpirationDate);

                switch (response.ResponseStatus)
                {
                    case EAuthResponseStatus.InvalidSignature:
                    case EAuthResponseStatus.InvalidResponseXML:
                    case EAuthResponseStatus.AuthenticationFailed:
                        TempData[TempDataKeys.EAuthErrorMessage] = response.ResponseStatusMessage;
                        return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "103", isIisError = true });

                    case EAuthResponseStatus.MissingEGN:
                        TempData[TempDataKeys.EAuthErrorMessage] = response.ResponseStatusMessage;
                        return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "101", isIisError = true });

                    case EAuthResponseStatus.NotDetectedQES:
                        TempData[TempDataKeys.EAuthErrorMessage] = response.ResponseStatusMessage;
                        return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "102", isIisError = true });

                    case EAuthResponseStatus.CanceledByUser:
                        return RedirectToAction(MVC.Account.ActionNames.Login, MVC.Account.Name);

                    case EAuthResponseStatus.Success:
                        {
                            InternalAdminUser internalAdminUser = this.systemRepository.GetInternalAdminUserByEgn(response.Egn);

                            if (internalAdminUser != null && internalAdminUser.IsActive)
                            {
                                EPaymentsAdmUser admUser = new EPaymentsAdmUser()
                                {
                                    UserId = internalAdminUser.InternalAdminUserId,
                                    Uin = internalAdminUser.Egn,
                                    Name = internalAdminUser.Name,
                                    Permission = internalAdminUser.IsSuperadmin ?
                                    InternalAdminUserPermissionEnum.Modify | InternalAdminUserPermissionEnum.ViewReferences | InternalAdminUserPermissionEnum.DistributionReferences :
                                    internalAdminUser.Permissions
                                };

                                var task = SignInManager.SignInAsync(admUser, false, false);
                                task.Wait();

                                if (internalAdminUser.IsSuperadmin)
                                {
                                    return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
                                }

                                if ((internalAdminUser.Permissions & InternalAdminUserPermissionEnum.Modify) == InternalAdminUserPermissionEnum.Modify)
                                {
                                    return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
                                }
                                else if ((internalAdminUser.Permissions & InternalAdminUserPermissionEnum.ViewReferences) == InternalAdminUserPermissionEnum.ViewReferences)
                                {
                                    return RedirectToAction(MVC.PaymentRequest.ActionNames.List, MVC.PaymentRequest.Name);
                                }
                                else if((internalAdminUser.Permissions & InternalAdminUserPermissionEnum.DistributionReferences) == InternalAdminUserPermissionEnum.DistributionReferences)
                                {
                                    return RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
                                }
                                else
                                {
                                    return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "104", isIisError = true });
                                }
                            }
                            else
                            {
                                return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "104", isIisError = true });
                            }
                        }
                    default:
                        throw new ArgumentException();
                }
            }
            else
            {
                return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "101", isIisError = true });
            }
        }
    }
}