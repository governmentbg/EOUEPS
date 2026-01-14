using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Admin.Models.ObligationType
{
    public class ObligationTypeVM
    {
        public List<SelectListItem> ObligationTypeSelectList { get; set; }

        public FormMode Mode { get; set; }

        public int ObligationTypeId { get; set; }

        [Required(ErrorMessage = "Полето „Име на типa задължение“ е задължително.")]
        [StringLength(200, ErrorMessage = "Дължината на „ Типa задължение“ не трябва да надвишава 200 символа.")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Полето „Статус на типа задължението“ е задължително.")]
        public ActiveStatus? IsActive { get; set; }

        public ObligationTypeVM()
        {
        }

        public ObligationTypeVM(FormMode mode, Model.Models.ObligationType obligation, List<EPayments.Model.Models.ObligationType> obligations)
        {
            this.Mode = mode;

            if (obligation != null)
            {
                this.ObligationTypeId = obligation.ObligationTypeId;
                this.Name = obligation.Name;
                this.IsActive = obligation.IsActive ? ActiveStatus.Active :ActiveStatus.InActive;
            }

            this.ObligationTypeSelectList = new List<SelectListItem>();
            this.ObligationTypeSelectList.Add(new SelectListItem() { Value = "", Text = "--Избери--" });
            this.ObligationTypeSelectList.AddRange(
                obligations.Select(e => new SelectListItem
                {
                    Value = e.ObligationTypeId.ToString(),
                    Text = e.Name
                }));
        }
    }

    public enum FormMode
    {
        Create,
        View,
        Edit,
    }

    public enum ActiveStatus
    {
        [Description("Активен")]
        Active = 0,

        [Description("Неактивен")]
        InActive = 1,
    }
}