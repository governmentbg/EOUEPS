using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace EPayments.Admin.Models.AdminUser
{
    public class InternalAdminUserVM
    {
        public List<SelectListItem> DepartmentSelectList { get; set; }

        public FormMode Mode { get; set; }

        public int InternalAdminUserId { get; set; }

        [Required(ErrorMessage = "Полето „Име“ е задължително.")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Полето „ЕГН“ е задължително.")]
        public string Egn { get; set; }

        [Required(ErrorMessage = "Полето „Суперадмин“ е задължително.")]
        public BoolNom? IsSuperadminBoolNom { get; set; }

        [Required(ErrorMessage = "Полето „Статус активност“ е задължително.")]
        public ActiveStatus? IsActiveBoolNom { get; set; }

        public PermissionEnum? Permission { get; set; }

        public ICollection<PermissionEnum> Permissions { get; set; }

        public InternalAdminUserVM()
        {
        }

        public InternalAdminUserVM(FormMode mode,  Model.Models.InternalAdminUser user)
        {
            this.Mode = mode;

            if (user != null)
            {
                this.InternalAdminUserId = user.InternalAdminUserId;
                this.Name = user.Name;
                this.Egn = user.Egn;
                this.IsSuperadminBoolNom = user.IsSuperadmin ? BoolNom.Yes : BoolNom.No;
                this.IsActiveBoolNom = user.IsActive ? ActiveStatus.Activated : ActiveStatus.Deactivated;

                if (user.Permissions != null)
                {
                    this.Permission = (PermissionEnum)(int)user.Permissions;
                }
            }
        }
    }

    public enum FormMode
    {
        Create,
        View,
        Edit,
    } 

    public enum BoolNom
    {
        [Description("Да")]
        Yes = 1,

        [Description("Не")]
        No = 2,
    }

    public enum ActiveStatus
    {
        [Description("Активен")]
        Activated = 1,

        [Description("Неактивен")]
        Deactivated = 2,
    }

    [Flags]
    public enum PermissionEnum
    {
        [Description("Преглед на справки")]
        ViewReferences = 1,
        [Description("Преглед на справки за разпределение")]
        DistributionReferences = 2,
        [Description("Управление")]
        Modify = 4
    }
}