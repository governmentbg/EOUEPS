using EPayments.Admin.Auth;
using EPayments.Admin.Common;
using EPayments.Admin.Controllers.DataObjects;
using EPayments.Admin.Models.Department;
using EPayments.Admin.Models.EserviceClient;
using EPayments.Admin.Models.Shared;
using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.Modify)]
    public partial class DepartmentController : BaseController
    {
        public const string AisKey = "AisClients";
        private const string UniqueIdentificationNumberExistsError = "Полето „ЕИК“ вече съществува в друга администрация.";
        private const string DepartmentNameExistsError = "Съществува друга администрация със същото име.";
        private const string NotFoundError = "Административната структура не беше намерена.";

        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;

        public DepartmentController(IWebRepository webRepository, IUnitOfWork unitOfWork)
        {
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult List(DepartmentListSearchDO searchDO)
        {
            DepartmentListVM model = new DepartmentListVM();

            var currentAdminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == this.CurrentUser.UserId);

            model.IsSuperadmin = currentAdminUser.IsSuperadmin;

            model.DepartmentRecordsPagingOptions = new PagingVM();

            model.SearchDO = searchDO;

            model.DepartmentRecordsPagingOptions.CurrentPageIndex = searchDO.Page;
            model.DepartmentRecordsPagingOptions.ControllerName = MVC.Department.Name;
            model.DepartmentRecordsPagingOptions.ActionName = MVC.Department.ActionNames.List;
            model.DepartmentRecordsPagingOptions.PageIndexParameterName = "page";
            model.DepartmentRecordsPagingOptions.RouteValues = searchDO.ToRouteValues();

            model.DepartmentRecordsPagingOptions.TotalItemCount = this.webRepository.CountDepartmentRecords(
                searchDO.DepartmentName,
                searchDO.DepartmentUniqueIdentificationNumber,
                searchDO.DepartmentUnifiedBudgetClassifier);

            model.DepartmentRecords = this.webRepository.GetDepartmentRecords(
                searchDO.DepartmentName,
                searchDO.DepartmentUniqueIdentificationNumber,
                searchDO.DepartmentUnifiedBudgetClassifier,
                searchDO.SortBy,
                searchDO.SortDesc,
                searchDO.Page,
                model.DepartmentRecordsPagingOptions.PageSize);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult ListSearch(DepartmentListSearchDO searchDO)
        {
            searchDO.Page = 1;

            return RedirectToListAction(searchDO);
        }

        [NonAction]
        private ActionResult RedirectToListAction(DepartmentListSearchDO searchDO)
        {
            return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name,
                new
                {
                    @departmentName = searchDO.DepartmentName,
                    @departmentUniqueIdentificationNumber = searchDO.DepartmentUniqueIdentificationNumber,
                    @departmentUnifiedBudgetClassifier = searchDO.DepartmentUnifiedBudgetClassifier,
                    @page = searchDO.Page,
                    @sortBy = searchDO.SortBy,
                    @sortDesc = searchDO.SortDesc,
                });
        }

        [HttpGet]
        public virtual ActionResult View(int id)
        {
            var department = this.webRepository.GetDepartment(id);

            if (department == null)
            {
                TempData[TempDataKeys.ErrorMessage] = "Административната структура не беше намерена.";

                return this.RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
            }

            DepartmentVM model = new DepartmentVM(Models.Department.FormMode.View, department);
            this.AddClients(model);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult Create()
        {
            DepartmentVM model = new DepartmentVM(Models.Department.FormMode.Create, null);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Create(DepartmentVM model)
        {
            var departments = this.unitOfWork.DbContext.Set<Department>().ToList();

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            if (this.unitOfWork.DbContext.Set<Department>().Any(d => d.Name == model.DepartmentName))
            {
                ModelState.AddModelError(string.Empty, DepartmentNameExistsError);
                return View(model);
            }

            if (!string.IsNullOrWhiteSpace(model.DepartmentUniqueIdentificationNumber) && this.unitOfWork.DbContext.Set<Department>().Any(d => d.UniqueIdentificationNumber == model.DepartmentUniqueIdentificationNumber))
            {
                ModelState.AddModelError(string.Empty, UniqueIdentificationNumberExistsError);
                return View(model);
            }

            var department = new Department();

            department.Name = model.DepartmentName;
            department.UniqueIdentificationNumber = model.DepartmentUniqueIdentificationNumber;
            department.UnifiedBudgetClassifier = model.DepartmentUnifiedBudgetClassifier;
            department.IsActive = true;

            this.unitOfWork.DbContext.Set<Department>().Add(department);
            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Административната структура беше успешно добавена.";

            return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
        }

        [HttpPost]
        public virtual ActionResult Delete(int id)
        {
            try
            {
                var department = this.unitOfWork.DbContext.Set<Department>()
                    .Include(d => d.EserviceClients)
                    .SingleOrDefault(d => d.DepartmentId == id);

                if (department == null)
                {
                    TempData[TempDataKeys.ErrorMessage] = NotFoundError;

                    return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
                }

                if (department.EserviceClients.Count > 0)
                {
                    TempData[TempDataKeys.ErrorMessage] = "Административната структура има АИС клиенти и не може да бъде изтрита.";

                    return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
                }

                this.unitOfWork.DbContext.Set<Department>().Remove(department);
                this.unitOfWork.Save();

                TempData[TempDataKeys.Message] = "Административната структура беше успешно изтритa.";

                return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
            }
            catch
            {
                TempData[TempDataKeys.ErrorMessage] = "Административната структура не беше изтритa.";
            }

            return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
        }

        [HttpGet]
        public virtual ActionResult Edit(int id)
        {
            var department = this.webRepository.GetDepartment(id);

            DepartmentVM model = new DepartmentVM(Models.Department.FormMode.Edit, department);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Edit(int id, DepartmentVM model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            if (this.unitOfWork.DbContext.Set<Department>().Any(d => d.Name == model.DepartmentName && d.DepartmentId != model.DepartmentId))
            {
                ModelState.AddModelError(string.Empty, DepartmentNameExistsError);
                return View(model);
            }

            if (!string.IsNullOrWhiteSpace(model.DepartmentUniqueIdentificationNumber) && this.unitOfWork.DbContext.Set<Department>().Any(d => d.UniqueIdentificationNumber == model.DepartmentUniqueIdentificationNumber && d.DepartmentId != model.DepartmentId))
            {
                ModelState.AddModelError(string.Empty, UniqueIdentificationNumberExistsError);
                return View(model);
            }

            var department = this.unitOfWork.DbContext.Set<Department>()
                .SingleOrDefault(d => d.DepartmentId == model.DepartmentId);

            if (department == null)
            {
                TempData[TempDataKeys.ErrorMessage] = NotFoundError;

                return RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
            }

            department.Name = model.DepartmentName;
            department.UniqueIdentificationNumber = model.DepartmentUniqueIdentificationNumber;
            department.UnifiedBudgetClassifier = model.DepartmentUnifiedBudgetClassifier;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за Административната структура бяха успешно редактирани.";
            return RedirectToAction(MVC.Department.ActionNames.View, MVC.Department.Name, new { id = id });
        }

        [HttpPost]
        public virtual ActionResult HierarchicalStructure(HierarchicalFormVM model)
        {
            Department department = this.unitOfWork.DbContext.Set<Department>()
                    .SingleOrDefault(d => d.EserviceClients.Any(ec => ec.EserviceClientId == model.EserviceClinetId));

            List<EserviceClient> clients = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Include(ec => ec.Children)
                .Where(ec => ec.EserviceClientId == model.EserviceClinetId || ec.EserviceClientId == model.TargetId)
                .ToList();

            if (department == null)
            {
                TempData[TempDataKeys.Message] = NotFoundError;
                return this.RedirectToAction(MVC.Department.ActionNames.List, MVC.Department.Name);
            }

            if (!this.ModelState.IsValid)
            {
                DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                this.AddClients(departmentModel);

                return this.View(MVC.Department.Views.View, departmentModel);
            }

            EserviceClient eserviceClient = department.EserviceClients.FirstOrDefault(ec => ec.EserviceClientId == model.EserviceClinetId);

            if (eserviceClient == null)
            {
                TempData[TempDataKeys.ErrorMessage] = "АИС клиента не е намерен.";

                DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                this.AddClients(departmentModel);

                return this.View(MVC.Department.Views.View, departmentModel);
            }

            if (eserviceClient.Children.Count > 0)
            {
                TempData[TempDataKeys.ErrorMessage] = "Не може да се промени разпоредителя на АИС клиента, защото той е разпоредител.";

                DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                this.AddClients(departmentModel);

                return this.View(MVC.Department.Views.View, departmentModel);
            }

            if (eserviceClient.AggregateToParent == true && model.TargetId == null)
            {
                TempData[TempDataKeys.ErrorMessage] = "Първостепеният разпоредител на АИС клиента не може да се премахне, защото АИС клиента агрегира към по-високо ниво.";

                DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                this.AddClients(departmentModel);

                return this.View(MVC.Department.Views.View, departmentModel);
            }

            EserviceClient targetClient = null;

            if (model.TargetId != null)
            {
                targetClient = clients.FirstOrDefault(ec => ec.EserviceClientId == model.TargetId);

                if (targetClient == null)
                {
                    TempData[TempDataKeys.ErrorMessage] = "Разпоредителя не е намерен.";

                    DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                    this.AddClients(departmentModel);

                    return this.View(MVC.Department.Views.View, departmentModel);
                }

                if (targetClient.DepartmentId != eserviceClient.DepartmentId)
                {
                    TempData[TempDataKeys.ErrorMessage] = "Разпоредителя не е в същата административна структура.";

                    DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                    this.AddClients(departmentModel);

                    return this.View(MVC.Department.Views.View, departmentModel);
                }

                if (targetClient.IsActive == false)
                {
                    TempData[TempDataKeys.ErrorMessage] = "Не можете да добавите за разпоредител, неактивен АИС клиент.";

                    DepartmentVM departmentModel = new DepartmentVM(Models.Department.FormMode.View, department);
                    this.AddClients(departmentModel);

                    return this.View(MVC.Department.Views.View, departmentModel);
                }
            }

            eserviceClient.ParentId = model.TargetId;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = string.Format("Разпоредителя на АИС клиент {0} беше променен на {1}.", 
                eserviceClient.AisName,
                targetClient?.AisName ?? "първостепенен разпоредител");

            return this.RedirectToAction(MVC.Department.Views.ViewNames.View, MVC.Department.Name, new { id = department.DepartmentId });
        }

        private DepartmentVM AddClients(DepartmentVM model)
        {
            model.AddClients(this.unitOfWork.DbContext.Set<EserviceClient>()
                .Where(ec => ec.DepartmentId == model.DepartmentId && ec.IsActive)
                .Select(EserviceVM.Map)
                .ToList());

            return model;
        }
    }
}