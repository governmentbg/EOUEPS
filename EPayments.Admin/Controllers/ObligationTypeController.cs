using EPayments.Admin.Auth;
using EPayments.Admin.Common;
using EPayments.Admin.Controllers.DataObjects;
using EPayments.Admin.Models.ObligationType;
using EPayments.Admin.Models.Shared;
using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.Modify)]
    public partial class ObligationTypeController : BaseController
    {
        //public const string ObligationTypesName = "ObligationTypes";

        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;
        public ObligationTypeController(IWebRepository webRepository, IUnitOfWork unitOfWork)
        {
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult List(ListObligationTypeSearchDO searchDO)
        {

            ObligationTypeLIstVM model = new ObligationTypeLIstVM();

            var currentAdminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == this.CurrentUser.UserId);

            model.IsSuperadmin = currentAdminUser.IsSuperadmin;

            model.ObligationTypeRecordsPagingOptions = new PagingVM();

            model.SearchDO = searchDO;
            model.ObligationTypeRecordsPagingOptions.CurrentPageIndex = searchDO.Page;
            model.ObligationTypeRecordsPagingOptions.ControllerName = MVC.ObligationType.Name;
            model.ObligationTypeRecordsPagingOptions.ActionName = MVC.ObligationType.ActionNames.List;
            model.ObligationTypeRecordsPagingOptions.PageIndexParameterName = "page";
            model.ObligationTypeRecordsPagingOptions.RouteValues = searchDO.ToRouteValues();

            model.ObligationTypeRecordsPagingOptions.TotalItemCount = this.webRepository.CountObligationTypeRecords(
                searchDO.Name,
                searchDO.IsActiveId);

            model.ObligationTypeRecords = this.webRepository.GetObligationTypeRecords(
                searchDO.Name,
                searchDO.IsActiveId,
                searchDO.SortBy,
                searchDO.SortDesc,
                searchDO.Page,
                10);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult View(int id)
        {
            var obligationType = this.webRepository.GetObligationType(id);
            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>().ToList();

            ObligationTypeVM model = new ObligationTypeVM(Models.ObligationType.FormMode.View, obligationType, obligationTypes);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult Create()
        {
            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>().ToList();

            ObligationTypeVM model = new ObligationTypeVM(FormMode.Create, null, obligationTypes);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Create(ObligationTypeVM model)
        {
            if (!ModelState.IsValid)
            {
                model.Mode = FormMode.Create;

                return View(model);
            }

            if (this.unitOfWork.DbContext.Set<ObligationType>().Any(ot => ot.Name == model.Name))
            {
                ModelState.AddModelError(string.Empty, "Съществува друг тип задължение със същото име.");
                
                model.Mode = FormMode.Create;

                return View(model);
            }

            var obligationType = new ObligationType();
            
            obligationType.Name = model.Name;
            obligationType.IsActive = model.IsActive.Value == ActiveStatus.Active;

            this.unitOfWork.DbContext.Set<ObligationType>().Add(obligationType);
            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Типа задължение беше успешно добавено.";

            return RedirectToAction(MVC.ObligationType.ActionNames.List, MVC.ObligationType.Name);
        }

        [HttpGet]
        public virtual ActionResult Edit(int id)
        {
            var obligationType = this.webRepository.GetObligationType(id);
            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>().ToList();

            ObligationTypeVM model = new ObligationTypeVM(FormMode.Edit, obligationType, obligationTypes);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Edit(int id, ObligationTypeVM model)
        {
            if (!ModelState.IsValid)
            {
                model.Mode = FormMode.Edit;
                
                return View(model);
            }

            if (this.unitOfWork.DbContext.Set<ObligationType>().Any(ot => ot.Name == model.Name && ot.ObligationTypeId != id))
            {
                ModelState.AddModelError(string.Empty, "Съществува друг тип задължение със същото име.");

                model.Mode = FormMode.Create;

                return View(model);
            }

            //if (model.IsActive == ActiveStatus.InActive)
            //{
            //    if (this.unitOfWork.DbContext.Set<EserviceClient>().Any(ec => ec.ObligationTypeId == id))
            //    {
            //        ModelState.AddModelError(string.Empty, "Не можете да деактивирате задължение, което се използва от АИС клиенти.");

            //        model.Mode = FormMode.Create;

            //        return View(model);
            //    }
            //}

            ObligationType obligationType = this.webRepository.GetObligationType(id);

            obligationType.Name = model.Name;
            obligationType.IsActive = model.IsActive == ActiveStatus.Active;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за типа на задължението бяха успешно редактирани.";
            return RedirectToAction(MVC.ObligationType.ActionNames.View, MVC.ObligationType.Name, new { id = id });
        }
    }
}