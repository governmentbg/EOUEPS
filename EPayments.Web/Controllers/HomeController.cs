using System;
using System.Web.Mvc;
using EPayments.Common.Captcha;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Models;
using EPayments.Web.Common;
using EPayments.Web.Models.Home;
using System.Linq;
using System.IO;
using EPayments.Common;
using EPayments.Model.DataObjects.EmailTemplateContext;
using System.Web;
using Microsoft.AspNet.Identity.Owin;
using System.Collections.Generic;
using EPayments.Web.DataObjects;
using System.Web.Helpers;
using IdentityModel.Client;
using System.Web.Hosting;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Text;
using EPayments.Web.Common.HelpDesk;
using EPayments.Web.VposHelpers.FiBank;
using EPayments.Web.VposHelpers.Dsk;
using System.Threading;

namespace EPayments.Web.Controllers
{
    public partial class HomeController : BaseController
    {
        private ISystemRepository systemRepository;
        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;

        private EPaymentsSignInManager SignInManager
        {
            get
            {
                return HttpContext.GetOwinContext().Get<EPaymentsSignInManager>();
            }
        }

        public HomeController(ISystemRepository systemRepository, IWebRepository webRepository, IUnitOfWork unitOfWork)
        {
            this.systemRepository = systemRepository;
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult Index()
        {
            if (Request.IsAuthenticated )
            {
                if (this.CurrentUser.EserviceAdminId.HasValue)
                {
                    return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
                }
                else
                {
                    return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                }
            }

            return View();
        }

        [HttpGet]
        public virtual ActionResult BanksInfo()
        {
            return View();
        }

        [HttpGet]
        public virtual ActionResult MdtMunicipality()
        {
            return View();
        }

        [HttpGet]
        public virtual ActionResult Help(string focus)
        {
            return View((object)focus);
        }

        [HttpGet]
        public virtual ActionResult Departments()
        {
            return View();
        }

        [HttpGet]
        public virtual ActionResult Feedback()
        {
            FeedbackVM model = new FeedbackVM();

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult Developers()
        {
            return View();
        }

        [HttpPost]
        [CaptchaValidation(Constants.CaptchaModelName)]
        public virtual ActionResult Feedback(FeedbackVM model, bool? captchaValid)
        {
            if (captchaValid.HasValue && !captchaValid.Value)
            {
                ModelState.AddModelError(Constants.CaptchaModelName, "Невалиден код за сигурност.");
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            string subjectText = String.Empty;
            switch (model.MessageType)
            {
                case "1":
                    subjectText = "Въпрос";
                    break;
                case "2":
                    subjectText = "Предложение";
                    break;
                case "3":
                    subjectText = "Технически проблем";
                    break;
                default:
                    break;
            }

            FeedbackContextDO contextDO = new FeedbackContextDO()
            {
                Name = model.Name ?? String.Empty,
                Email = model.Email ?? String.Empty,
                Phone = model.Phone ?? String.Empty,
                Subject = subjectText,
                MessageText = model.Message == null ? String.Empty : Formatter.ConvertTextToHtml(model.Message)
            };

            Email email = new Email(contextDO, AppSettings.EPaymentsWeb_FeedbackEmail);

            this.systemRepository.AddEntity<Email>(email);
            this.unitOfWork.Save();

            TempData[TempDataKeys.FeedbackSend] = true;

            if (AppSettings.EPaymentsWeb_HelpDeskSubmitEnabled)
            {
                Task.Run(() =>
                {
                    SendFeedbackToHelpDesk(subjectText, model.Name, model.Email, model.Phone, model.Message);
                });
            }

            return RedirectToAction(MVC.Home.ActionNames.Feedback, MVC.Home.Name);
        }

        [HttpGet]
        public virtual ActionResult AccessByCode(string code)
        {
            AccessByCodeVM model = new AccessByCodeVM();
            model.AccessCode = code;

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult AccessByPaymentRequestCode(string code)
        {
            AccessByCodeVM model = new AccessByCodeVM();
            model.AccessCode = code;

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestByAccessCode(model.AccessCode);
            if (paymentRequest == null)
            {
                ModelState.AddModelError(String.Empty, "Невалиден код за плащане.");

                return View("AccessByCode", model);
            }
            else
            {
                EPaymentsUser ePaymentsUser = new EPaymentsUser()
                {
                    IsAuthorizedByAccessCode = true,
                    AccessCode = paymentRequest.PaymentRequestAccessCode
                };

                var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                task.Wait();

                return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
        }

        [HttpPost]
        [CaptchaValidation(Constants.CaptchaModelName)]
        public virtual ActionResult AccessByCode(AccessByCodeVM model, bool? captchaValid)
        {
            if (captchaValid.HasValue && !captchaValid.Value)
            {
                ModelState.AddModelError(Constants.CaptchaModelName, "Невалиден код за сигурност.");
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestByAccessCode(model.AccessCode);
            if (paymentRequest == null)
            {
                ModelState.AddModelError(String.Empty, "Невалиден код за плащане.");

                return View(model);
            }
            else
            {
                EPaymentsUser ePaymentsUser = new EPaymentsUser()
                {
                    IsAuthorizedByAccessCode = true,
                    AccessCode = paymentRequest.PaymentRequestAccessCode
                };

                var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                task.Wait();

                return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
        }

        [HttpGet]
        public virtual ActionResult AccessByEserviceAdmin()
        {
            if (Request.IsAuthenticated)
            {
                if (this.CurrentUser.EserviceAdminId.HasValue)
                {
                    return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);
                }
            }

            AccessByEserviceAdminVM model = new AccessByEserviceAdminVM();

            return View(model);
        }

        [HttpPost]
        [CaptchaValidation(Constants.CaptchaModelName)]
        public virtual ActionResult AccessByEserviceAdmin(AccessByEserviceAdminVM model, bool? captchaValid)
        {
            if (captchaValid.HasValue && !captchaValid.Value)
            {
                ModelState.AddModelError(Constants.CaptchaModelName, "Невалиден код за сигурност.");
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            bool isValidLogin = false;

            EserviceAdminUser eserviceAdminUser = this.systemRepository.GetActiveEserviceAdminUsername(model.Username);

            if (eserviceAdminUser != null)
            {
                //string salt = Crypto.GenerateSalt();
                //string hash = Crypto.HashPassword(password + salt);

                isValidLogin = Crypto.VerifyHashedPassword(eserviceAdminUser.PasswordHash, model.Password + eserviceAdminUser.PasswordSalt);
            }

            if (!isValidLogin)
            {
                ModelState.AddModelError(String.Empty, "Невалидно потребителско име или парола.");

                return View(model);
            }
            else
            {
                if (!String.IsNullOrWhiteSpace(eserviceAdminUser.IpAddressesForAccess))
                {
                    List<string> allowedIpAddresses = eserviceAdminUser.IpAddressesForAccess.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                    string ipAddress = this.Request.UserHostAddress;

                    if (!allowedIpAddresses.Any(e => e.Trim() == ipAddress))
                    {
                        //Access from this IP address is not allowed
                        throw new Exception();

                        //ModelState.AddModelError(String.Empty, "Нямате право на достъп от този IP адрес.");
                        //return View(model);
                    }
                }

                EPaymentsUser ePaymentsUser = new EPaymentsUser()
                {
                    Name = eserviceAdminUser.Name,
                    Uin = null,
                    UserId = null,
                    EserviceAdminId = eserviceAdminUser.EserviceAdminUserId,
                    //EserviceAdminDepartment = eserviceAdminUser.Department.Name
                };

                var task = SignInManager.SignInAsync(ePaymentsUser, false, false);
                task.Wait();

                return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
            }
        }

        [HttpGet]
        public virtual ActionResult AccessibilityPolicy()
        {
            return View();
        }

        [HttpGet]
        public virtual ActionResult AccessRules()
        {
            return View();
        }

        [HttpGet]
        public virtual ActionResult About()
        {
            return View();
        }

        [HttpGet]
        public virtual ActionResult SystemStats(SystemStatsSearchDO searchDO)
        {
            SystemStatsVM model = new SystemStatsVM();

            //Init DropdownValues
            model.PeriodYearDropdownValues = new Dictionary<string, string>();
            model.PeriodYearDropdownValues.Add(String.Empty, "--Всички--");
            for(int i = DateTime.Now.Year; i >= 2015; i--)
            {
                model.PeriodYearDropdownValues.Add(i.ToString(), i.ToString());
            }

            List<EserviceClient> allEserviceClients = this.systemRepository.GetAllEserviceClients();
            model.EserviceClientIdDropdownValues = new Dictionary<string,string>();
            model.EserviceClientIdDropdownValues.Add(String.Empty, "--Всички--");
            foreach (var eservicClient in allEserviceClients)
            {
                model.EserviceClientIdDropdownValues.Add(eservicClient.EserviceClientId.ToString(), eservicClient.AisName);
            }

            List<Department> allDepartments = this.systemRepository.GetAllDepartments();
            model.DepartmentIdDropdownValues = new Dictionary<string,string>();
            model.DepartmentIdDropdownValues.Add(String.Empty, "--Всички--");
            foreach (var department in allDepartments)
            {
                model.DepartmentIdDropdownValues.Add(department.DepartmentId.ToString(), department.Name);
            }

            //Get stats results
            model.SystemStats = this.systemRepository.GetSystemStats(searchDO.PeriodYear, searchDO.DepartmentId, searchDO.EserviceClientId);

            model.SearchDO = searchDO;

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult SystemStatsSearch(SystemStatsSearchDO searchDO)
        {
            return RedirectToAction(MVC.Home.ActionNames.SystemStats, MVC.Home.Name,
                new
                {
                    @periodYear = searchDO.PeriodYear,
                    @departmentId = searchDO.DepartmentId,
                    @eserviceClientId = searchDO.EserviceClientId
                });
        }

        [HttpGet]
        public virtual FileResult DownloadFile(string id)
        {
            string localFileName = null;
            string saveFileName = null;
            string mimeType = null;

            switch (id.ToLower().Trim())
            {
                case "ruleselectronicservicepdf":
                    {
                        localFileName = "Общи условия за присъединяване.pdf";
                        saveFileName = "Общи условия за присъединяване на административни органи и на доставчици на платежни услуги към Единната входна точка за електронни плащания в централната и местната администрация и използване на функционалността 'Централен ВПОС терминал', предоставяна от Министерство на електронното управление.pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                //case "rulespaymentservicepdf":
                //    {
                //        localFileName = "rulesPaymentService.pdf";
                //        saveFileName = "Правила за включване на доставчици на платежни услуги.pdf";
                //        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                //        break;
                //    }
                case "apppaymentservicevpospdf":
                    {
                        localFileName = "Prilojenie-1.1-DPU.pdf";
                        saveFileName = "Заявление за присъединяване на доставчик на платежни услуги чрез използване на Виртуален ПОС терминал към Единната входна точка за електронни плащания в централната и местната администрация(Единната входна точка).pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                case "apppaymentservicecvpospdf":
                    {
                        localFileName = "Prilojenie-1.2-DPU-CVPOS.pdf";
                        saveFileName = "Заявление за присъединяване на доставчик на платежни услуги към Единната входна точка за електронни плащания в централната и местната администрация(Единната входна точка) чрез използване на функционалността 'Централен виртуален ПОС терминал', предоставяна от Министерство на електронното управление (МЕУ).pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                //case "appelectronicservicedoc":
                //    {
                //        localFileName = "appElectronicService.doc";
                //        saveFileName = "Заявление за включване на доставчик на електронни административни услуги.doc";
                //        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                //        break;
                //    }
                case "appelectronicservicepdf":
                    {
                        localFileName = "Prilojenie-2.2-AO-spr.pdf";
                        saveFileName = "Заявление за присъединяване на административен орган към Единната входна точка за електронни плащания в централната и местната администрация с цел предоставяне на електронната административна услуга 'Извличане на справка за дължими данъци и други публичноправни задължения и плащането им по електронен път'.pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                case "appelectronicservicecvpospdf":
                    {
                        localFileName = "Prilojenie-2.1-AO-CVPOS.pdf";
                        saveFileName = "Заявление за присъединяване /промяна на обстоятелства/ на административен орган към Единната входна точка за електронни плащания в централната и местната администрация, чрез виртуален ПОС терминал на доставчик на платежни услуги и/ или чрез използване на функционалността 'Централен виртуален ПОС терминал', предоставяна от Министерство на електронното управление (МЕУ).pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                
                case "appelectronicserviceterminationpdf":
                    {
                        localFileName = "Prilojenie-3.2-AO-prekr.pdf";
                        saveFileName = "Заявление за прекратяване присъединяването на административен орган към Единната входна точка за електронни плащания в централната и местната администрация(Единната входна точка).pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                case "apppaymentserviceterminationpdf":
                    {
                        localFileName = "Prilojenie-3.1-DPU-prekr.pdf";
                        saveFileName = "Заявление за прекратяване присъединяването на доставчик на платежни услуги към Единната входна точка за електронни плащания в централната и местната администрация (Единната входна точка).pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                case "procedureborikapdf":
                    {
                        localFileName = "Prawila_БОРИКА osporvane.pdf";
                        saveFileName = "ПРОЦЕДУРА ПО 'ОСПОРВАНЕ НА ПЛАЩАНЕ', ИЗВЪРШЕНО ЧРЕЗ ВИРТУАЛНОТО ТЕРМИНАЛНО ПОС УСТРОЙСТВО ПРЕДОСТАВЕНО НА МИНИСТЕРСТВО НА ЕЛЕКТРОННОТО УПРАВЛЕНИЕ, ОБСЛУЖВАНО ОТ 'БОРИКА' АД.pdf";
                        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                //case "apppaymentservicedoc":
                //    {
                //        localFileName = "appPaymentService.doc";
                //        saveFileName = "ЗАЯВЛЕНИЕ за присъединяване на доставчик на платежни услуги към Средата за електронни плащания чрез използване на Централен Виртуален ПОС терминал.pdf";
                //        mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                //        break;
                //    }
                default:
                    throw new ArgumentException();
            }

            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", localFileName);

            byte[] data = System.IO.File.ReadAllBytes(filePath);

            return File(data, mimeType, saveFileName);
        }

        [HttpGet]
        public virtual ActionResult RedirectToError(string id, int? logId, string egn, bool? isIisError, string url)
        {
            TempData[TempDataKeys.ErrorId] = id;
            TempData[TempDataKeys.ErrorAttemptLogId] = logId;
            TempData[TempDataKeys.ErrorUserEgn] = egn;
            TempData[TempDataKeys.IsIisError] = isIisError;
            TempData[TempDataKeys.Url] = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, url).ToString();

            return RedirectToAction(MVC.Home.ActionNames.Error, MVC.Home.Name);
        }

        [HttpGet]
        public virtual ActionResult Error()
        {
            if (TempData[TempDataKeys.ErrorId] == null)
            {
                return RedirectToAction(MVC.Home.ActionNames.Index, MVC.Home.Name);
            }

            CustomHandleErrorInfo errorInfo = new CustomHandleErrorInfo(new Exception());
            errorInfo.ErrorCode = (string)TempData[TempDataKeys.ErrorId];
            errorInfo.AttempLogId = (int?)TempData[TempDataKeys.ErrorAttemptLogId];
            errorInfo.Egn = (string)TempData[TempDataKeys.ErrorUserEgn];
            errorInfo.Url = (string)TempData[TempDataKeys.Url];
            errorInfo.EAuthError = (string)TempData[TempDataKeys.EAuthErrorMessage];

            bool? isIisError = (bool?)TempData[TempDataKeys.IsIisError];
            if (isIisError.HasValue && isIisError.Value)
            {
                LoginAttemptLog attemptLog = new LoginAttemptLog();
                attemptLog.IP = Formatter.TruncateString(Request.UserHostAddress, 50);
                attemptLog.LogDate = DateTime.Now;
                attemptLog.ErrorCode = errorInfo.ErrorCode;
                attemptLog.IsIisErrorOccurred = true;
                attemptLog.IsUesParsed = false;
                attemptLog.IsLoginSuccessful = false;

                this.systemRepository.AddEntity<LoginAttemptLog>(attemptLog);

                this.unitOfWork.Save();
            }

            return View(MVC.Shared.Views._Error, errorInfo);
        }

        [HttpGet]
        public virtual ActionResult EidLogin()
        {
            return View();
        }

        [NonAction]
        private void SendFeedbackToHelpDesk(string subjectText, string name, string email, string phone, string desctiption)
        {
            string accessToken = null;

            using (TokenClient tokenClient = new TokenClient(
                Formatter.UriCombine(AppSettings.EPaymentsWeb_HelpDeskUrl, "api/token").ToString(),
                AppSettings.EPaymentsWeb_HelpDeskClientId,
                AppSettings.EPaymentsWeb_HelpDeskClientSecret))
            {
                var tokenResponse = tokenClient.RequestResourceOwnerPasswordAsync(AppSettings.EPaymentsWeb_HelpDeskUsername, AppSettings.EPaymentsWeb_HelpDeskPassword);
                var tokenResponseResult = tokenResponse.Result;

                if (tokenResponseResult.IsError || tokenResponseResult.TokenType != "Bearer")
                    throw new Exception(String.Format("HelpDesk get access token exception - isError: {0}, tokenType: {1}",
                        tokenResponseResult.IsError,
                        tokenResponseResult.TokenType));

                accessToken = String.Format("Bearer {0}", tokenResponseResult.AccessToken);
            }

            object postData = new 
            {
                name = String.Format("Съобщение за: '{0}'{1}{2}",
                    subjectText,
                    !String.IsNullOrWhiteSpace(name) ? String.Format(" от {0}", name.Trim()) : String.Empty,
                    !String.IsNullOrWhiteSpace(phone) ? String.Format(" ({0})", phone.Trim()) : String.Empty),
                description = !String.IsNullOrWhiteSpace(desctiption) ? desctiption.Trim() : null,
                issuerEmail = !String.IsNullOrWhiteSpace(email) ? email.Trim() : null
            };

            string postDataJson = JsonConvert.SerializeObject(postData);

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(AppSettings.EPaymentsWeb_HelpDeskUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Add("Authorization", accessToken);

                StringContent postContent = new StringContent(postDataJson, Encoding.UTF8, "application/json");

                var responseResult = client.PostAsync("/api/units/createSimple", postContent).Result;

                if (!responseResult.IsSuccessStatusCode)
                    throw new Exception(String.Format("HelpDesk create unit exception - statusCode: {0}", responseResult.StatusCode));

                ActionResultDO actionResultDO = JsonConvert.DeserializeObject<ActionResultDO>(responseResult.Content.ReadAsStringAsync().Result);

                if (actionResultDO.ResultCode != ActionResultCode.Ok)
                    throw new Exception(String.Format("HelpDesk create unit exception - actionResultDO.ResultCode: {0}", actionResultDO.ResultCode));
            }
        }
    }
}