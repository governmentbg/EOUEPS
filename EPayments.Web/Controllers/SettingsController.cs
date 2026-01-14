using System;
using System.Web.Mvc;
using EPayments.Model.Models;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Common.Data;
using EPayments.Web.Models.Settings;
using EPayments.Web.Auth;
using EPayments.Web.Models.Shared;
using EPayments.Common;
using EPayments.Web.DataObjects;

namespace EPayments.Web.Controllers
{
    [WebUserAuthorize]
    public partial class SettingsController : BaseController
    {
        private ISystemRepository systemRepository;
        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;

        public SettingsController(ISystemRepository systemRepository, IWebRepository webRepository, IUnitOfWork unitOfWork)
        {
            this.systemRepository = systemRepository;
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult Index()
        {
            return RedirectToAction(MVC.Settings.ActionNames.View, MVC.Settings.Name);
        }

        [HttpGet]
        public virtual ActionResult Edit()
        {
            User user = this.systemRepository.GetUserById(this.CurrentUser.UserId.Value);

            EditVM model = new EditVM();
            model.Email = user.Email;
            model.RequestNotifications = user.RequestNotifications;
            model.StatusNotifications = user.StatusNotifications;
            model.AccessCodeNotifications = user.AccessCodeNotifications;
            model.StatusObligationNotifications = user.StatusObligationNotifications;

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Edit(EditVM model)
        {
            if ((model.RequestNotifications || model.StatusNotifications || model.AccessCodeNotifications) && String.IsNullOrWhiteSpace(model.Email))
            {
                ModelState.AddModelError("Email", "Полето „Електронна поща“ е задължително.");
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            User user = this.systemRepository.GetUserById(this.CurrentUser.UserId.Value);
            user.Email = !String.IsNullOrWhiteSpace(model.Email) ? model.Email : null;
            user.RequestNotifications = model.RequestNotifications;
            user.StatusNotifications = model.StatusNotifications;
            user.StatusObligationNotifications = model.StatusObligationNotifications;
            user.AccessCodeNotifications = model.AccessCodeNotifications;

            this.unitOfWork.Save();

            return RedirectToAction(MVC.Settings.ActionNames.View, MVC.Settings.Name);
        }

        [HttpGet]
        public virtual ActionResult View(RequestAccessListSearchDO searchDO)
        {
            User user = this.systemRepository.GetUserById(this.CurrentUser.UserId.Value);

            ViewVM model = new ViewVM();

            model.Email = user.Email;
            model.RequestNotifications = user.RequestNotifications;
            model.StatusNotifications = user.StatusNotifications;
            model.AccessCodeNotifications = user.AccessCodeNotifications;
            model.StatusObligationNotifications = user.StatusObligationNotifications;
            model.SearchDO = searchDO;

            model.RequestAccessListPagingOptions = new PagingVM();
            model.RequestAccessListPagingOptions.CurrentPageIndex = searchDO.Page;
            model.RequestAccessListPagingOptions.ControllerName = MVC.Settings.Name;
            model.RequestAccessListPagingOptions.ActionName = MVC.Settings.ActionNames.View;
            model.RequestAccessListPagingOptions.PageIndexParameterName = "page";
            model.RequestAccessListPagingOptions.RouteValues = searchDO.ToRequestAccessListRouteValues();

            model.RequestAccessListPagingOptions.TotalItemCount = this.webRepository.CountRequestAccessByUin(this.CurrentUser.Uin);

            model.RequestAccessList = this.webRepository.GetRequestAccessByUin(
                this.CurrentUser.Uin,
                searchDO.SortBy,
                searchDO.SortDesc,
                searchDO.Page,
                AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult ViewSort(RequestAccessListSearchDO searchDO)
        {
            return RedirectToAction(MVC.Settings.ActionNames.View, MVC.Settings.Name,
                new
                {
                    @page = searchDO.Page,
                    @sortBy = searchDO.SortBy,
                    @sortDesc = searchDO.SortDesc,
                });
        }

        [HttpGet]
        public virtual ActionResult RequestAccessDetails(Guid id)
        {
            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);

            RequestAccessDetailsVM model = new RequestAccessDetailsVM();
            model.PaymentRequestIdentifier = paymentRequest.PaymentRequestIdentifier;
            model.LimitCount = 50;
            model.AccessCount = this.webRepository.GetRequestAccessDetailsCount(paymentRequest.PaymentRequestId, this.CurrentUser.Uin);
            model.AccessDetails.AddRange(this.webRepository.GetRequestAccessDetails(paymentRequest.PaymentRequestId, this.CurrentUser.Uin, model.LimitCount));

            return PartialView(model);
        }
    }
}