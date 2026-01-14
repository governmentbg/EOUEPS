using EPayments.Admin.Auth;
using EPayments.Admin.Common;
using EPayments.Admin.DataObjects;
using EPayments.Admin.Models.AdminUser;
using EPayments.Admin.Models.Shared;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.Modify)]
    public partial class AdminUserController : BaseController
    {
        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;

        public AdminUserController(IWebRepository webRepository, IUnitOfWork unitOfWork)
        {
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult List()
        {
            CheckSuperadminPersmissions();

            var model = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .OrderByDescending(e => e.CreateDate)
                .ToList();

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult View(int id)
        {
            CheckSuperadminPersmissions();

            var adminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == id);

            InternalAdminUserVM model = new InternalAdminUserVM(FormMode.View, adminUser);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult Create()
        {
            CheckSuperadminPersmissions();

            InternalAdminUserVM model = new InternalAdminUserVM(FormMode.Create, null);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Create(InternalAdminUserVM model)
        {
            CheckSuperadminPersmissions();

            if (!string.IsNullOrWhiteSpace(model.Egn))
            {
                EgnHelper egnObj = new EgnHelper(model.Egn);
                if (!egnObj.IsValid())
                {
                    ModelState.AddModelError(string.Empty, "Стойността за „ЕГН“ е невалидна.");
                }
                else if (this.unitOfWork.DbContext.Set<InternalAdminUser>().Any(e => e.Egn == model.Egn))
                {
                    ModelState.AddModelError(string.Empty, "Вече съществува добавен администратор с това ЕГН.");
                }
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var adminUser = new InternalAdminUser();

            adminUser.Name = model.Name;
            adminUser.Egn = model.Egn;
            adminUser.IsSuperadmin = model.IsSuperadminBoolNom.Value == BoolNom.Yes;
            adminUser.IsActive = model.IsActiveBoolNom == ActiveStatus.Activated;
            adminUser.CreateDate = DateTime.Now;

            if (model.Permissions != null && model.Permissions.Count > 0)
            {
                adminUser.Permissions = (InternalAdminUserPermissionEnum)(int)model.Permissions.Aggregate((acc, curr) => acc | curr);
            }

            this.unitOfWork.DbContext.Set<InternalAdminUser>().Add(adminUser);
            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Администратора беше успешно добавен.";

            return RedirectToAction(MVC.AdminUser.ActionNames.List, MVC.AdminUser.Name);
        }

        [HttpGet]
        public virtual ActionResult Edit(int id)
        {
            CheckSuperadminPersmissions();

            var adminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == id);

            InternalAdminUserVM model = new InternalAdminUserVM(FormMode.Edit, adminUser);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Edit(int id, InternalAdminUserVM model)
        {
            CheckSuperadminPersmissions();

            if (model.Permissions != null && model.Permissions.Count > 0)
            {
                model.Permission = (PermissionEnum?)(int)model.Permissions.Aggregate((acc, curr) => acc | curr);
            }

            var adminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == id);

            if (!string.IsNullOrWhiteSpace(model.Egn))
            {
                EgnHelper egnObj = new EgnHelper(model.Egn);
                if (!egnObj.IsValid())
                {
                    ModelState.AddModelError(string.Empty, "Стойността за „ЕГН“ е невалидна.");
                }
                else if (this.unitOfWork.DbContext.Set<InternalAdminUser>().Any(e => e.Egn == model.Egn && e.InternalAdminUserId != adminUser.InternalAdminUserId))
                {
                    ModelState.AddModelError(string.Empty, "Вече съществува администратор с това ЕГН.");
                }
            }

            bool isAdminActive = model.IsActiveBoolNom == ActiveStatus.Activated;
            bool isSuperAdmin = model.IsSuperadminBoolNom.Value == BoolNom.Yes;

            if (adminUser.InternalAdminUserId == this.CurrentUser.UserId && adminUser.IsActive != isAdminActive)
            {
                ModelState.AddModelError(string.Empty, "Не можете да промените \"Статус активност\" на своя потребител.");
            }

            if (adminUser.InternalAdminUserId == this.CurrentUser.UserId && adminUser.IsSuperadmin != isSuperAdmin)
            {
                ModelState.AddModelError(string.Empty, "Не можете да редактирате ролята на суперадмин на своя потребител.");
            }

            if (adminUser.InternalAdminUserId == this.CurrentUser.UserId && (int?)model.Permission != (int?)adminUser.Permissions)
            {
                ModelState.AddModelError(string.Empty, "Не можете да променяте правата на своя потребител.");
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            adminUser.Name = model.Name;
            adminUser.Egn = model.Egn;
            adminUser.IsSuperadmin = isSuperAdmin;
            adminUser.IsActive = isAdminActive;

            if (model.Permissions != null && model.Permissions.Count > 0)
            {
                adminUser.Permissions = (InternalAdminUserPermissionEnum)(int)model.Permissions.Aggregate((acc, curr) => acc | curr);
            }

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за администратора бяха успешно редактирани.";

            return RedirectToAction(MVC.AdminUser.ActionNames.View, MVC.AdminUser.Name, new { id = id });
        }

        private void CheckSuperadminPersmissions()
        {
            var currentAdminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == this.CurrentUser.UserId);

            if (!currentAdminUser.IsSuperadmin && (currentAdminUser.Permissions & InternalAdminUserPermissionEnum.Modify) == 0)
                throw new Exception("Unauthorized action. User is not superadmin.");
        }
    }
}