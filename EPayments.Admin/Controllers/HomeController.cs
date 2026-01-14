using EPayments.Admin.Common;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Models;
using System;
using System.Configuration;
using System.IO;
using System.Web.Configuration;
using System.Web.Mvc;

namespace EPayments.Admin.Controllers
{
    public partial class HomeController : BaseController
    {
        private ISystemRepository systemRepository;
        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;

        public HomeController(ISystemRepository systemRepository, IWebRepository webRepository, IUnitOfWork unitOfWork)
        {
            this.systemRepository = systemRepository;
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult RedirectToError(string id, int? logId, string egn, bool? isIisError, string url)
        {
            TempData[TempDataKeys.ErrorId] = id;
            TempData[TempDataKeys.ErrorAttemptLogId] = logId;
            TempData[TempDataKeys.ErrorUserEgn] = egn;
            TempData[TempDataKeys.IsIisError] = isIisError;
            TempData[TempDataKeys.Url] = url;

            return RedirectToAction(MVC.Home.ActionNames.Error, MVC.Home.Name);
        }

        [HttpGet]
        public virtual ActionResult Error()
        {
            if (TempData[TempDataKeys.ErrorId] != null)
            {
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
                    attemptLog.ErrorCode = $"Epayments.Admin error: {errorInfo.ErrorCode}";
                    attemptLog.IsIisErrorOccurred = true;
                    attemptLog.IsUesParsed = false;
                    attemptLog.IsLoginSuccessful = false;

                    this.systemRepository.AddEntity<LoginAttemptLog>(attemptLog);

                    this.unitOfWork.Save();
                }

                return View(MVC.Shared.Views._Error, errorInfo);
            }
            else
            {
                return View(MVC.Shared.Views._Error, null);
            }
        }
    }
}