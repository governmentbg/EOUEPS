using System;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using EPayments.Web.Common;
using EPayments.CertificateUtils.Extractor;
using EPayments.CertificateUtils;
using EPayments.Model.Models;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Common.Data;
using EPayments.Web.Models.Account;
using System.Xml;
using System.Text;
using EPayments.Web.Models.Shared;
using EPayments.Web.Auth;
using EPayments.Common.DataObjects;
using Newtonsoft.Json.Linq;
using EPayments.Common;
using JWT;
using JWT.Serializers;
using EPayments.Common.Saml;
using System.Security.Cryptography;
using System.IO;
using SAML2.Schema.Metadata;
using SAML2.Schema.XmlDSig;
using System.Security.Cryptography.Xml;

namespace EPayments.Web.Controllers
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

        private EPaymentsSignInManager SignInManager
        {
            get
            {
                return HttpContext.GetOwinContext().Get<EPaymentsSignInManager>();
            }
        }

        private ICertificateExtractor certificateExtractor;
        private ISystemRepository systemRepository;
        private IUnitOfWork unitOfWork;

        public AccountController(ICertificateExtractor certificateExtractor, ISystemRepository systemRepository, IUnitOfWork unitOfWork)
        {
            this.certificateExtractor = certificateExtractor;
            this.systemRepository = systemRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        [AllowAnonymous]
        public ActionResult Metadata()
        {
            var pathToCertificate = AppSettings.EPaymentsWeb_EAuthRequestSignCertificatePath;
            if (string.IsNullOrEmpty(pathToCertificate))
            {
                throw new ArgumentNullException("pathToCertificate");
            }

            if (!Path.IsPathRooted(pathToCertificate))
            {
                var dirname = System.Web.HttpContext.Current.Server.MapPath("~/");
                pathToCertificate = Path.Combine(dirname, pathToCertificate);
            }
            if (!System.IO.File.Exists(pathToCertificate))
            {
                throw new ArgumentOutOfRangeException("pathToCertificate");
            }

            // Load your certificate.
            X509Certificate2 x509Certificate = new X509Certificate2(pathToCertificate, AppSettings.EPaymentsWeb_EAuthRequestSignCertificatePass);

            var id = Guid.NewGuid().ToString();
            // Create Entity Descriptor with ID received from the IdP.
            EntityDescriptor descriptor = new EntityDescriptor();
            descriptor.EntityID = Dns.GetHostName();
            descriptor.ValidUntil = DateTime.Now;
            descriptor.ID = id;
            SpSsoDescriptor spd = new SpSsoDescriptor();
            spd.Id = Dns.GetHostName();
            spd.ProtocolSupportEnumeration = new[] { "urn:oasis:names:tc:SAML:2.0:protocol" };
            spd.AuthnRequestsSigned = "true";
            spd.WantAssertionsSigned = "true";

            // Creating a key descriptor.
            KeyDescriptor keyDescriptor = new KeyDescriptor();
            keyDescriptor.UseSpecified = true;
            keyDescriptor.Use = KeyTypes.Signing;
            X509Data keyData = new X509Data();
            keyData.Items = new object[] { x509Certificate.RawData };
            keyData.ItemsElementName = new X509ItemType[] { X509ItemType.X509Certificate };

            // Create KeyInfo.
            SAML2.Schema.XmlDSig.KeyInfo keyInfo = new SAML2.Schema.XmlDSig.KeyInfo();
            keyInfo.Items = new object[] { keyData };
            keyInfo.ItemsElementName = new KeyInfoItemType[] { KeyInfoItemType.X509Data };
            keyDescriptor.KeyInfo = keyInfo;


            // Creating second key descriptor.
            KeyDescriptor keyDescriptor2 = new KeyDescriptor();
            keyDescriptor2.UseSpecified = true;
            keyDescriptor2.Use =  KeyTypes.Encryption;
            keyDescriptor2.KeyInfo = keyInfo;

            // Add KeyDescriptor.
            spd.KeyDescriptor = new[] { keyDescriptor, keyDescriptor2 };

            string returnUri = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Account.Name, MVC.Account.ActionNames.EAuthResponse).ToString();
            // Assign assertion service URL.
            IndexedEndpoint consumerService = new IndexedEndpoint();
            consumerService.Index = 0;
            consumerService.IsDefault = true;
            consumerService.Location = returnUri;
            consumerService.Binding = "urn:oasis:names:tc:SAML:2.0:bindings:HTTP-POST";
            spd.AssertionConsumerService = new IndexedEndpoint[] { consumerService };


            IndexedEndpoint artifactResolutionService = new IndexedEndpoint();
            artifactResolutionService.Index = 0;
            artifactResolutionService.Location = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Account.Name, MVC.Account.ActionNames.Logout).ToString();
            artifactResolutionService.Binding = "urn:oasis:names:tc:SAML:2.0:bindings:HTTP-GET";

            IndexedEndpoint artifactResolutionService2 = new IndexedEndpoint();
            artifactResolutionService2.Index = 0;
            artifactResolutionService2.Location = returnUri;
            artifactResolutionService2.Binding = "urn:oasis:names:tc:SAML:2.0:bindings:HTTP-POST";

            spd.ArtifactResolutionService = new IndexedEndpoint[] { artifactResolutionService, artifactResolutionService2 };

            Endpoint singleLogoutService = new Endpoint();
            singleLogoutService.Binding = "urn:oasis:names:tc:SAML:2.0:bindings:HTTP-GET";
            singleLogoutService.Location = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Account.Name, MVC.Account.ActionNames.Logout).ToString();
            singleLogoutService.ResponseLocation = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Account.Name, MVC.Account.ActionNames.Logout).ToString();
            spd.SingleLogoutService = new Endpoint[] { singleLogoutService };

            descriptor.Items = new object[] { spd };

            // Add some information.
            // Organization information
            descriptor.Organization = new Organization();
            descriptor.Organization.OrganizationName = new LocalizedName[] { new LocalizedName("DAEU", "en") };
            descriptor.Organization.OrganizationDisplayName = new LocalizedName[] { new LocalizedName("DAEU", "en") };
            descriptor.Organization.OrganizationURL = new LocalizedURI[] { new LocalizedURI("https://e-gov.bg", "en") };

            // Add contact person info.
            Contact person = new Contact();
            person.Company = "DAEU";
            person.EmailAddress = new[] { "some@some.bg" };
            // Contact information
            descriptor.ContactPerson = new Contact[] { person };
            // Sign metadata with service provider key
            SAML2.Schema.XmlDSig.Signature signature = new SAML2.Schema.XmlDSig.Signature();
            signature.KeyInfo = keyInfo;
            var signatureValue = new SignatureValue();
            signatureValue.Value = x509Certificate.PublicKey.EncodedKeyValue.RawData;

            signature.SignatureValue = signatureValue;

            var signatureInfo = new SAML2.Schema.XmlDSig.SignedInfo();

            var canonicalizationMethod = new CanonicalizationMethod();
            canonicalizationMethod.Algorithm = SignedXml.XmlDsigExcC14NTransformUrl;
            signatureInfo.CanonicalizationMethod = canonicalizationMethod;

            //var signatureMethod = new SignatureMethod();
            //signatureMethod.Algorithm = System.Security.Cryptography.Xml.SignedXml.XmlDsigRSASHA1Url;
            //signatureInfo.SignatureMethod = signatureMethod;

            SAML2.Schema.XmlDSig.Reference reference = new SAML2.Schema.XmlDSig.Reference();
            var digestMethod = new DigestMethod();
            digestMethod.Algorithm = SignedXml.XmlDsigSHA256Url;
            reference.DigestMethod = digestMethod;
            reference.DigestValue = x509Certificate.GetCertHash();
            reference.URI = "#" + id;
            var t1 = new SAML2.Schema.XmlDSig.Transform();
            t1.Algorithm = SignedXml.XmlDsigEnvelopedSignatureTransformUrl;

            var t2 = new SAML2.Schema.XmlDSig.Transform();
            t2.Algorithm = SignedXml.XmlDsigExcC14NTransformUrl;

            reference.Transforms = new SAML2.Schema.XmlDSig.Transform[] { t1, t2 };

            signatureInfo.Reference = new SAML2.Schema.XmlDSig.Reference[] { reference };
            signature.SignedInfo = signatureInfo;

            descriptor.Signature = signature;

            string xmlString = SAML2.Utils.Serialization.SerializeToXmlString(descriptor);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlString);

            XmlElement elem = doc.CreateElement("SignatureMethod");
            elem.SetAttribute("Algorithm", SignedXml.XmlDsigRSASHA256Url);
            doc.DocumentElement.GetElementsByTagName("Signature")[0].FirstChild.InsertBefore(elem, doc.GetElementsByTagName("Reference")[0]);

            return this.Content(doc.OuterXml, "text/xml");
        }

        [HttpGet]
        [AllowAnonymous]
        public virtual async Task<ActionResult> Login()
        {
            if (Request.IsAuthenticated)
            {
                return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);
            }

            byte[] certificateRawContent = (byte[])Session[SessionKeys.CertificateRawContent];
            Session.Remove(SessionKeys.CertificateRawContent);

            X509Certificate2 x509Certificate = new X509Certificate2(certificateRawContent);

            CertificateDO certificateDO = this.certificateExtractor.ExtractCertificateDO(x509Certificate);

            LoginAttemptLog attemptLog = ValidateCertificateAndLogAttempt(certificateRawContent, x509Certificate, certificateDO);

            if (attemptLog.IsLoginSuccessful)
            {
                EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, false);

                await SignInManager.SignInAsync(ePaymentsUser, false, false);

                return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);
            }
            else
            {
                return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = attemptLog.ErrorCode, logId = attemptLog.LoginAttemptLogId });
            }
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult Logout()
        {
            AuthenticationManager.SignOut();

            return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);
        }

        [HttpGet]
        [EserviceAdminAuthorize]
        public virtual ActionResult LogoutEserviceAdminUser()
        {
            AuthenticationManager.SignOut();

            return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);
        }

        [HttpGet]
        [AllowAnonymous]
        public virtual ActionResult EAuth()
        {
            if (AppSettings.EPaymentsWeb_EAuthEnabled)
            {
                string returnUri = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Account.Name, MVC.Account.ActionNames.EAuthResponse).ToString();
                string metadataUri = AppSettings.EPaymentsWeb_EAuthMetadataUrl;

                XmlDocument authnRequest = SamlHelper.GenerateKEPAuthnRequest(
                    AppSettings.EPaymentsWeb_PortalName,
                    AppSettings.EPaymentsWeb_EAuthProviderId,
                    returnUri,
                    AppSettings.EPaymentsWeb_EAuthRequestUrl,
                    AppSettings.EPaymentsWeb_EAuthExtServiceId,
                    AppSettings.EPaymentsWeb_EAuthExtProviderId,
                    metadataUri);

                string signedAuthnRequest = string.Empty;
                signedAuthnRequest = SamlHelper.SignXmlDocument(authnRequest, AppSettings.EPaymentsWeb_EAuthRequestSignCertificatePath, AppSettings.EPaymentsWeb_EAuthRequestSignCertificatePass);

                var base64AuthnRequest = Convert.ToBase64String(Encoding.UTF8.GetBytes(signedAuthnRequest));

                AutoPostVM model = new AutoPostVM();
                model.MethodType = "post";
                model.PostUrl = AppSettings.EPaymentsWeb_EAuthRequestUrl;
                model.PostValues = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>>();
                model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("SAMLRequest", base64AuthnRequest));
                model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("RelayState", String.Empty));

                return View(MVC.Shared.Views._AutoPost, model);
            }
            else
            {
                return Redirect("~/SSL/Cert.aspx");
            }
        }

        [HttpPost]
        [AllowAnonymous]
        public virtual ActionResult EAuthResponse(string samlResponse)
        {
            if (AppSettings.EPaymentsWeb_EAuthEnabled)
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
                        case EAuthResponseStatus.MissingEGN:
                        case EAuthResponseStatus.NotDetectedQES:
                            TempData[TempDataKeys.EAuthErrorMessage] = response.ResponseStatusMessage;
                            return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "101", isIisError = true });

                        case EAuthResponseStatus.CanceledByUser:
                            return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);

                        case EAuthResponseStatus.Success:
                            CertificateDO certificateDO = new CertificateDO()
                            {
                                PersonalIdentifier = response.Egn,
                                Name = response.Name,
                                IsPersonal = true,
                                IsCompany = false,
                                X509Certificate = null
                            };

                            EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, true);

                            var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                            task.Wait();

                            return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);

                        default:
                            throw new ArgumentException();
                    }
                }
                else
                {
                    return RedirectToAction(MVC.Home.ActionNames.RedirectToError, MVC.Home.Name, new { id = "101", isIisError = true });
                }
            }
            else
            {
                return Redirect("~/SSL/Cert.aspx");
            }
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult AuthPassLogin(AuthRequestDO requestDO)
        {
            if (AppSettings.EPaymentsWeb_IsAuthPassEnabled)
            {
                var client = this.systemRepository.GetEserviceClientByClientId(requestDO.ClientId);
                if (client.IsAuthPassAuthorized)
                {
                    string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                    JObject dataJObject = JObject.Parse(dataJson);

                    string egn = dataJObject.GetValue("egn", StringComparison.OrdinalIgnoreCase).Value<string>();
                    string name = dataJObject.GetValue("name", StringComparison.OrdinalIgnoreCase).Value<string>();
                    string certBase64 = dataJObject.GetValue("cert", StringComparison.OrdinalIgnoreCase).Value<string>();
                    string guid = dataJObject.GetValue("guid", StringComparison.OrdinalIgnoreCase).Value<string>();

                    if (!this.systemRepository.IsAuthPassLoginExist(Guid.Parse(guid)))
                    {
                        AuthPassLogin authPassLogin = new AuthPassLogin();
                        authPassLogin.Gid = Guid.Parse(guid);
                        authPassLogin.IP = this.Request.UserHostAddress;
                        authPassLogin.PostData = dataJson;
                        authPassLogin.LogDate = DateTime.Now;

                        this.systemRepository.AddEntity<AuthPassLogin>(authPassLogin);
                        this.unitOfWork.Save();

                        CertificateDO certificateDO = new CertificateDO()
                        {
                            PersonalIdentifier = egn,
                            Name = name,
                            IsPersonal = true,
                            IsCompany = false,
                            X509Certificate = null
                        };

                        EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, true);

                        var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                        task.Wait();

                        return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                    }
                    else
                    {
                        throw new Exception("AuthPass login is corrupted.");
                    }
                }
                else
                {
                    throw new Exception("Eservice client is not authorized for AuthPass login.");
                }
            }
            else
            {
                throw new Exception("AuthPass login is disabled.");
            }
        }

        [HttpGet]
        [AllowAnonymous]
        public virtual ActionResult EidSubmitForm()
        {
            var model = new EidVM();

            return View(MVC.Shared.Views._EidSubmitForm, model);
        }

        [AllowAnonymous]
        [HttpGet]
        public virtual ActionResult Eid(string Target, string URL, string SAMLArtifact)
        {
            if (AppSettings.EPaymentsWeb_EidEnabled)
            {
                if (!string.IsNullOrEmpty(SAMLArtifact))
                {
                    try
                    {
                        var artifactResolveMessage = SamlHelper.GenerateSaml2ArtifactResolve(SAMLArtifact);
                        //ElmahLogger.Instance.Info(String.Format("URL: {0}, ArtefactMessage: {1}", URL, artifactResolveMessage));
                        var response = SamlHelper.SendWebRequest(artifactResolveMessage, URL);
                        //ElmahLogger.Instance.Info("Response: " + response);
                        SAML2.Schema.Protocol.Status status;

                        DcElectronicIdentityInfo eIdResponse = SamlHelper.ParseSaml2Result(response, out status);

                        if (status.StatusCode.Value != SamlHelper.eIDSuccessStatus)
                        {
                            throw new Exception("Invalid EID Status");
                            //ElmahLogger.Instance.Info("eid status response: " + status.StatusCode.Value);
                            //ModelState.AddModelError("eID", status.StatusMessage);
                            //return View("Login", model);
                        }

                        if (eIdResponse == null)
                        {
                            throw new Exception("Invalid EID Response");
                            //ElmahLogger.Instance.Info("invalid eid response: " + response);
                            //ModelState.AddModelError("eID", ErrorMessages.InvalidEIDResponse);
                        }

                        CertificateDO certificateDO = new CertificateDO()
                        {
                            PersonalIdentifier = eIdResponse.EGN,
                            Name = Formatter.FormatName(eIdResponse.GivenNameLat, eIdResponse.MiddleNameLat, eIdResponse.FamilyNameLat),
                            IsPersonal = true,
                            IsCompany = false,
                            X509Certificate = null
                        };

                        EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, true);

                        var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                        task.Wait();

                        return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("EID exception " + ex.Message ?? "" + " innerMsg: " + (ex.InnerException != null ? ex.Message ?? "" : ""));
                    }
                }
                else
                {
                    throw new Exception("Invalid SAMLArtifact");
                }
            }
            else
            {
                throw new Exception("Eid is disabled.");
            }
        }

        [AllowAnonymous]
        [HttpGet]
        public virtual ActionResult NoiAuth(string jwt)
        {
            if (AppSettings.EPaymentsWeb_NoiAuthEnabled)
            {
                try
                {
                    JsonNetSerializer serializer = new JsonNetSerializer();
                    JwtBase64UrlEncoder urlEncoder = new JwtBase64UrlEncoder();

                    IDateTimeProvider dateTimeProvider = new UtcDateTimeProvider();
                    IJwtValidator validator = new JwtValidator(serializer, dateTimeProvider);

                    JwtDecoder decoder = new JwtDecoder(serializer, validator, urlEncoder);
                    string jwtJsonData = decoder.Decode(jwt, AppSettings.EPaymentsWeb_NoiAuthSecretKey, true);

                    JObject jwtJObject = JObject.Parse(jwtJsonData);

                    string egn = jwtJObject.GetValue("EGN", StringComparison.OrdinalIgnoreCase).Value<string>();
                    string firstName = jwtJObject.GetValue("FirstName", StringComparison.OrdinalIgnoreCase).Value<string>();
                    string secondName = jwtJObject.GetValue("Surname", StringComparison.OrdinalIgnoreCase).Value<string>();
                    string lastName = jwtJObject.GetValue("LastName", StringComparison.OrdinalIgnoreCase).Value<string>();

                    CertificateDO certificateDO = new CertificateDO()
                    {
                        PersonalIdentifier = egn,
                        Name = Formatter.FormatName(firstName, secondName, lastName),
                        IsPersonal = true,
                        IsCompany = false,
                        X509Certificate = null
                    };

                    EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, true);

                    var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                    task.Wait();

                    return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                }
                catch
                {
                    throw new Exception("Invalid NoiAuth token");
                }
            }
            else
            {
                throw new Exception("NoiAuth is disabled.");
            }
        }

        [AllowAnonymous]
        [HttpGet]
        public virtual ActionResult EDeliveryAuth(string token)
        {
            if (AppSettings.EPaymentsWeb_EDeliveryAuthEnabled)
            {
                JsonNetSerializer serializer = new JsonNetSerializer();
                JwtBase64UrlEncoder urlEncoder = new JwtBase64UrlEncoder();

                IDateTimeProvider dateTimeProvider = new UtcDateTimeProvider();
                IJwtValidator validator = new JwtValidator(serializer, dateTimeProvider);

                JwtDecoder decoder = new JwtDecoder(serializer, validator, urlEncoder);
                string jwtJsonData = decoder.Decode(token, AppSettings.EPaymentsWeb_EDeliveryAuthSecret, true);

                JObject jwtJObject = JObject.Parse(jwtJsonData);

                string iss = jwtJObject.GetValue("iss", StringComparison.OrdinalIgnoreCase).Value<string>();
                string sub = jwtJObject.GetValue("sub", StringComparison.OrdinalIgnoreCase).Value<string>();

                if (iss.Trim().ToLower() != "urn:oid:2.16.100.1.1.1.1.13" || !sub.Trim().ToLower().StartsWith("urn:egn:"))
                    throw new Exception();

                string egnEncrypted = sub.Trim().Substring("urn:egn:".Length);

                string egn = DecryptEDeliveryUin(egnEncrypted, AppSettings.EPaymentsWeb_EDeliveryAuthSecret);

                EgnHelper egnObj = new EgnHelper(egn);
                if (!egnObj.IsValid())
                    throw new Exception("Invalid EGN");

                CertificateDO certificateDO = new CertificateDO()
                {
                    PersonalIdentifier = egn,
                    Name = null,
                    IsPersonal = true,
                    IsCompany = false,
                    X509Certificate = null
                };

                EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, true);

                var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                task.Wait();

                return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
            else
            {
                throw new Exception("EDeliveryAuth is disabled.");
            }
        }

        [AllowAnonymous]
        [HttpGet]
        public virtual ActionResult EMySpaceAuth(string token)
        {
            if (AppSettings.EPaymentsWeb_EMySpaceAuthEnabled)
            {
                JsonNetSerializer serializer = new JsonNetSerializer();
                JwtBase64UrlEncoder urlEncoder = new JwtBase64UrlEncoder();

                IDateTimeProvider dateTimeProvider = new UtcDateTimeProvider();
                IJwtValidator validator = new JwtValidator(serializer, dateTimeProvider);

                JwtDecoder decoder = new JwtDecoder(serializer, validator, urlEncoder);
                string jwtJsonData = decoder.Decode(token, AppSettings.EPaymentsWeb_EMySpaceAuthSecret, true);

                JObject jwtJObject = JObject.Parse(jwtJsonData);

                string iss = jwtJObject.GetValue("iss", StringComparison.OrdinalIgnoreCase).Value<string>();
                string sub = jwtJObject.GetValue("sub", StringComparison.OrdinalIgnoreCase).Value<string>();

                if (iss.Trim().ToLower() != string.Format("urn:oid:{0}", AppSettings.EPaymentsWeb_EMySpaceAuthOID) || !sub.Trim().ToLower().StartsWith("urn:egn:"))
                    throw new Exception();

                string egnEncrypted = sub.Trim().Substring("urn:egn:".Length);

                string egn = DecryptEMySpaceUin(egnEncrypted, AppSettings.EPaymentsWeb_EMySpaceAuthSecret);

                EgnHelper egnObj = new EgnHelper(egn);
                if (!egnObj.IsValid())
                    throw new Exception("Invalid EGN");

                CertificateDO certificateDO = new CertificateDO()
                {
                    PersonalIdentifier = egn,
                    Name = null,
                    IsPersonal = true,
                    IsCompany = false,
                    X509Certificate = null
                };

                EPaymentsUser ePaymentsUser = GetEPaymentsUser(certificateDO, true);

                var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                task.Wait();

                return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
            else
            {
                throw new Exception("EDeliveryAuth is disabled.");
            }
        }

        [HttpPost]
        [AllowAnonymous]
        public virtual ActionResult EAuthAdminRedirect(string samlResponse, string relayState)
        {
            AutoPostVM model = new AutoPostVM();
            model.PostUrl = AppSettings.EPaymentsWeb_EAuthAdminRedirect;
            model.PostValues = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>>();
            model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("samlResponse", samlResponse));
            model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("relayState", relayState));

            return View(MVC.Shared.Views._AutoPost, model);
        }

        [HttpPost]
        [AllowAnonymous]
        public virtual ActionResult EAuthAdminRedirectTest(string samlResponse, string relayState)
        {
            string redirectToAdminUrl = @"http://172.23.104.201:8899/Account/EAuthLogin?samlResponse={samlResponse}&relayState={relayState}";

            return Redirect(redirectToAdminUrl);
        }

        private string DecryptEDeliveryUin(string cipherText, string sharedSecret)
        {
            byte[] salt = new byte[] { 34, 134, 145, 12, 7, 6, 243, 63, 43, 54, 75, 65, 53, 2, 34, 54, 45, 67, 64, 64, 32, 213 };

            return Decrypt(cipherText, sharedSecret, salt);
        }

        private string DecryptEMySpaceUin(string cipherText, string sharedSecret)
        {
            byte[] salt = Convert.FromBase64String(AppSettings.EPaymentsWeb_EMySpaceAuthSecretSalt);

            return this.Decrypt(cipherText, sharedSecret, salt);
        }

        private string Decrypt(string cipherText, string sharedSecret, byte[] salt)
        {
            RijndaelManaged aesAlg = null;
            string result = null;

            try
            {
                Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, salt);
                byte[] bytes = Convert.FromBase64String(cipherText);
                using (MemoryStream msDecrypt = new MemoryStream(bytes))
                {
                    aesAlg = new RijndaelManaged();
                    aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                    aesAlg.IV = ReadByteArray(msDecrypt);
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                        {
                            result = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }
            finally
            {
                if (aesAlg != null)
                {
                    aesAlg.Clear();
                }
            }

            return result;
        }

        private byte[] ReadByteArray(Stream s)
        {
            byte[] rawLength = new byte[sizeof(int)];
            if (s.Read(rawLength, 0, rawLength.Length) != rawLength.Length)
            {
                throw new SystemException("Stream did not contain properly formatted byte array");
            }

            byte[] buffer = new byte[BitConverter.ToInt32(rawLength, 0)];
            if (s.Read(buffer, 0, buffer.Length) != buffer.Length)
            {
                throw new SystemException("Did not read byte array properly");
            }

            return buffer;
        }

        private LoginAttemptLog ValidateCertificateAndLogAttempt(byte[] certificateRawContent, X509Certificate2 x509Certificate, CertificateDO certificateDO)
        {
            LoginAttemptLog attemptLog = new LoginAttemptLog();
            attemptLog.LogDate = DateTime.Now;
            attemptLog.IP = Formatter.TruncateString(Request.UserHostAddress, 50);
            attemptLog.CertificateFile = certificateRawContent;
            attemptLog.CertificateIssuer = x509Certificate.Issuer;
            attemptLog.CertificatePolicies = x509Certificate.Extensions["Certificate Policies"] != null ? x509Certificate.Extensions["Certificate Policies"].Format(true) : null;
            attemptLog.CertificateSubject = x509Certificate.Subject;
            attemptLog.AlternativeSubject = x509Certificate.Extensions["Subject Alternative Name"] != null ? x509Certificate.Extensions["Subject Alternative Name"].Format(true) : null;
            attemptLog.CertificateThumbprint = x509Certificate.Thumbprint;
            attemptLog.Egn = certificateDO != null ? certificateDO.PersonalIdentifier : null;
            attemptLog.NameLatin = certificateDO != null ? certificateDO.Name : null;
            attemptLog.IsIisErrorOccurred = false;
            attemptLog.IsUesParsed = certificateDO != null;
            attemptLog.IsPersonal = certificateDO != null ? certificateDO.IsPersonal : (bool?)null;
            attemptLog.IsEgnParsed = certificateDO != null ? certificateDO.HasPersonalIdentifier : (bool?)null;
            attemptLog.IsNameParsed = certificateDO != null ? !String.IsNullOrWhiteSpace(certificateDO.Name) : (bool?)null;

            if (!AppSettings.EPaymentsWeb_UseFakeCertificate && DateTime.Now > x509Certificate.NotAfter || DateTime.Now < x509Certificate.NotBefore)
            {
                attemptLog.ErrorCode = "104";
                attemptLog.IsLoginSuccessful = false;
            }
            else if (certificateDO == null)
            {
                attemptLog.ErrorCode = "201";
                attemptLog.IsLoginSuccessful = false;
            }
            else if (!certificateDO.HasPersonalIdentifier)
            {
                attemptLog.ErrorCode = "203";
                attemptLog.IsLoginSuccessful = false;
            }
            else
            {
                attemptLog.IsLoginSuccessful = true;
            }

            #region Certificate type validation

            #endregion

            this.systemRepository.AddEntity<LoginAttemptLog>(attemptLog);

            this.unitOfWork.Save();

            return attemptLog;
        }

        private EPaymentsUser GetEPaymentsUser(CertificateDO certificateDO, bool isEAuthLogin)
        {
            Certificate certificate = null;
            if (!isEAuthLogin)
            {
                certificate = this.systemRepository.GetCertificateByThumbprint(certificateDO.X509Certificate.Thumbprint);

                if (certificate == null)
                {
                    certificate = new Certificate();
                    certificate.CertificateSubject = certificateDO.X509Certificate.Subject;
                    certificate.CertificateIssuer = certificateDO.X509Certificate.Issuer;
                    certificate.CertificateFile = certificateDO.X509Certificate.GetRawCertData();
                    certificate.CertificateThumbprint = certificateDO.X509Certificate.Thumbprint;

                    this.systemRepository.AddEntity<Certificate>(certificate);

                    this.unitOfWork.Save();
                }
            }

            User user = this.systemRepository.GetUserByUin(certificateDO.PersonalIdentifier);

            if (user == null)
            {
                user = new User();
                user.Egn = certificateDO.PersonalIdentifier;
                user.RequestNotifications = false;
                user.StatusNotifications = false;
                user.StatusObligationNotifications = false;

                this.systemRepository.AddEntity<User>(user);
            }

            if (!isEAuthLogin)
            {
                user.LastCertificateId = certificate.CertificateId;
            }

            this.unitOfWork.Save();

            return new EPaymentsUser()
            {
                Name = certificateDO.Name,
                Uin = certificateDO.PersonalIdentifier,
                UserId = user.UserId,
                EserviceAdminId = null,
                EserviceAdminDepartment = null
            };
        }
    }
}