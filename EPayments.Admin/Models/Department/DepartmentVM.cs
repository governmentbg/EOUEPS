using EPayments.Admin.Models.EserviceClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Admin.Models.Department
{
    public class DepartmentVM
    {
        public FormMode Mode { get; set; }

        public int DepartmentId { get; set; }

        [Required(ErrorMessage = "Полето „Име на нова администрация“ е задължително.")]
        [StringLength(200, ErrorMessage = "Дължината на „Администрация“ не трябва да надвишава 200 символа.")]
        public string DepartmentName { get; set; }

        [StringLength(16, ErrorMessage = "Дължината на „ЕИК“ трябва да бъде максимум 16 символа.")]
        [RegularExpression(@"^[0-9]*$", ErrorMessage = "Полето „ЕИК“ може да съдържа само цифри.")]
        [Required(ErrorMessage = "Полето „ЕИК“ е задължително.")]
        public string DepartmentUniqueIdentificationNumber { get; set; }

        [StringLength(16, ErrorMessage = "Дължината на „ЕБК“ трябва да бъде максимум 16 символа.")]
        [RegularExpression(@"^[0-9]*$", ErrorMessage = "Полето „ЕБК“ може да съдържа само цифри.")]
        public string DepartmentUnifiedBudgetClassifier { get; set; }

        public ICollection<EserviceVM> EserviceRecords { get; private set; } = new HashSet<EserviceVM>();

        public ICollection<SelectListItem> SelectListItemsOfEserviceClients { get; private set; } = new HashSet<SelectListItem>();

        public DepartmentVM()
        {
        }

        public DepartmentVM(FormMode mode, Model.Models.Department department)
        {
            this.Mode = mode;

            if (department != null)
            {
                this.DepartmentId = department.DepartmentId;
                this.DepartmentName = department.Name;
                this.DepartmentUniqueIdentificationNumber = department.UniqueIdentificationNumber;
                this.DepartmentUnifiedBudgetClassifier = department.UnifiedBudgetClassifier;
            }
        }

        public void AddClients(ICollection<EserviceVM> eserviceRecords)
        {
            if (eserviceRecords != null && eserviceRecords.Count > 0)
            {
                foreach (EserviceVM record in eserviceRecords)
                {
                    if (record.ParentId == null)
                    {
                        this.EserviceRecords.Add(record);
                    }
                    else
                    {
                        EserviceVM parent = eserviceRecords.FirstOrDefault(ec => ec.EserviceClientId == record.ParentId);

                        if (parent != null)
                        {
                            parent.Children.Add(record);
                        }
                    }

                    this.SelectListItemsOfEserviceClients.Add(new SelectListItem()
                    {
                        Text = record.AisName,
                        Value = record.EserviceClientId.ToString()
                    });
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
}