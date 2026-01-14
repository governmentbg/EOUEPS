using EPayments.Admin.Models.ObligationType;
using EPayments.Common.Helpers;
using EPayments.EDelivery.Models;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web.Mvc;
using Client = EPayments.Model.Models.EserviceClient;

namespace EPayments.Admin.Models.EserviceClient
{
    public class EServiceClientVM : IValidatableObject
    {
        public List<SelectListItem> DepartmentSelectList { get; set; } = new List<SelectListItem>();

        public List<SelectListItem> EServiceClientSelectList { get; set; } = new List<SelectListItem>();

        public FormMode Mode { get; set; }

        public int EserviceClientId { get; set; }

        [Required(ErrorMessage = "Полето „Име на АИС“ е задължително.")]
        [StringLength(100, ErrorMessage = "Дължината на „Име на АИС“ не трябва да надвишава 100 символа.")]
        public string AisName { get; set; }

        [Required(ErrorMessage = "Полето „Избор на администрация“ е задължително.")]
        public int DepartmentId { get; set; }

        public string DepartmentName { get; set; }

        [Required(ErrorMessage = "Полето „Обслужваща банка“ е задължително.")]
        [StringLength(100, ErrorMessage = "Дължината на „Обслужваща банка“ не трябва да надвишава 100 символа.")]
        public string AccountBank { get; set; }

        [Required(ErrorMessage = "Полето „BIC код на обслужваща банка“ е задължително.")]
        [RegularExpression(@"^[A-Za-z0-9]*$", ErrorMessage = "Полето „BIC код на обсл. банка“ може да съдържа само символи на латиница и цифри.")]
        [StringLength(8, MinimumLength = 8, ErrorMessage = "Дължината на „BIC код на обсл. банка“ трябва да бъде 8 символа.")]
        public string AccountBIC { get; set; }

        [Required(ErrorMessage = "Полето „IBAN на обслужваща сметка“ е задължително.")]
        [StringLength(22, MinimumLength = 22, ErrorMessage = "Дължината на „IBAN на обслужваща сметка“ трябва да бъде 22 символа.")]
        [RegularExpression(@"^[A-Za-z0-9]*$", ErrorMessage = "Полето „IBAN на обслужваща сметка“ може да съдържа само символи на латиница и цифри.")]
        public string AccountIBAN { get; set; }

        [Required(ErrorMessage = "Полето „Поддръжка на EPAY“ е задължително.")]
        public bool IsEpayVposEnabled { get; set; }

        [Required(ErrorMessage = "Полето „Поддръжка на Борика“ е задължително.")]
        public bool IsBoricaVposEnabled { get; set; }

        [Required(ErrorMessage = "Полето „Статус активност“ е задължително.")]
        public ActiveStatus? IsActiveBoolNom { get; set; }

        [Required(ErrorMessage = "Полето „Агрегация към по-високо ниво“ е задължително.")]
        public BoolNom AggregateToParent { get; set; }

        [NotEqualToProperty(nameof(EserviceClientId), ErrorMessage = "АИС клиента не може да си е родител в йерархията.")]
        public int? ParentId { get; set; }

        [Required(ErrorMessage = "Полето „Вид на разпределение“ е задължително.")]
        public DistributionType DistributionType { get; set; }

        //public int ObligationTypeId { get; set;}
        public string DeliveryAdminstrationId { get; set; }
        public string DeliveryAdministrationGuid { get; set; }

        //[StringLength(10, ErrorMessage = "Дължината на бюджетният код не може да е по-дълга от 10 символа.")]
        //public string BudgetCode { get; set; }

        public string ClientId { get; set; }
        public string SecretKey { get; set; }
        public Vpos? VposClientId { get; set; }

        public int? EserviceAdminId { get; set; }
        public string EserviceAdminUsername { get; set; }
        public string EserviceAdminName { get; set; }

        public List<SelectListItem> ObligationTypes { get; private set; }

        public List<SelectListItem> DeliveryAdministrationList { get; set; }
        public List<KeyValuePair<string, string>> DeliveryAdministrationUIList { get; set; }

        public EServiceClientVM()
        {
        }

        public EServiceClientVM(
            FormMode mode,  
            Client eserviceClient, 
            EserviceAdminUser eserviceAdminUser,
            List<EPayments.Model.Models.Department> departments,
            List<ObligationTypeOptionVM> obligationTypes,
            DepartmentInstitutionInfo[] deliveryAdministrations,
            List<Client> eServiceClients)
        {
            this.Mode = mode;

            if (eserviceClient != null)
            {
                this.EserviceClientId = eserviceClient.EserviceClientId;
                this.AisName = eserviceClient.AisName;
                this.DepartmentId = eserviceClient.DepartmentId;
                this.AccountBank = eserviceClient.AccountBank;
                this.AccountBIC = eserviceClient.AccountBIC;
                this.AccountIBAN = eserviceClient.AccountIBAN;
                this.ClientId = eserviceClient.ClientId;
                this.SecretKey = eserviceClient.SecretKey;
                this.VposClientId = (Vpos?)eserviceClient.VposClientId;
                this.IsEpayVposEnabled = eserviceClient.IsEpayVposEnabled;
                this.IsBoricaVposEnabled = eserviceClient.IsBoricaVposEnabled;
                this.IsActiveBoolNom = eserviceClient.IsActive ? ActiveStatus.Activated : ActiveStatus.Deactivated;
                //this.ObligationTypeId = eserviceClient.ObligationTypeId;
                this.DeliveryAdministrationGuid = eserviceClient.DeliveryAdministrationGuid;
                //this.BudgetCode = eserviceClient.BudgetCode;
                if (deliveryAdministrations != null)
                {
                    this.DeliveryAdminstrationId =
                   deliveryAdministrations
                       .FirstOrDefault(x => x.ElectronicSubjectId.ToString().Equals(eserviceClient.DeliveryAdministrationGuid, System.StringComparison.OrdinalIgnoreCase))?.UniqueSubjectIdentifier ?? null;
                }
                this.AggregateToParent = eserviceClient.AggregateToParent ? BoolNom.Yes : BoolNom.No;
                this.ParentId = eserviceClient.ParentId;

                if (eserviceClient.DistributionTypeId == 1)
                {
                    this.DistributionType = DistributionType.WithBankTransfer;
                }

            }

            this.SetServiceEClients(eServiceClients);

            this.SetDeparments(departments);

            //this.SetObligationTypes(obligationTypes);
            this.SetDeliveryAdminstrations(deliveryAdministrations);

            this.SetAdminUser(eserviceAdminUser);
        }

        //public void SetObligationTypes(List<ObligationTypeOptionVM> obligationTypes)
        //{
        //    if (obligationTypes != null && obligationTypes.Count > 0)
        //    {
        //        this.ObligationTypes = new List<SelectListItem>();
        //        this.ObligationTypes.AddRange(obligationTypes.Select(o => new SelectListItem() { Value = o.ObligationTypeId.ToString(), Text = o.Name, Selected = o.ObligationTypeId == this.ObligationTypeId }));
        //    }
        //}

        public void SetServiceEClients(List<Client> eServiceClients)
        {
            if (eServiceClients != null && eServiceClients.Count > 1)
            {
                this.EServiceClientSelectList = new List<SelectListItem>();

                this.EServiceClientSelectList.Add(new SelectListItem() { Value = "", Text = "--Избери--", Selected = this.ParentId == null });
                this.EServiceClientSelectList.AddRange(
                    eServiceClients.Where(e => this.EserviceClientId != e.EserviceClientId)
                    .Select(e => new SelectListItem
                    {
                        Value = e.EserviceClientId.ToString(),
                        Text = e.AisName,
                        Selected = (e.EserviceClientId == this.ParentId)
                    }));
            }
        }

        public void SetDeparments(List<EPayments.Model.Models.Department> departments)
        {
            if (departments != null && departments.Count > 0)
            {
                this.DepartmentSelectList.AddRange(
                departments.Select(e => new SelectListItem
                {
                    Value = e.DepartmentId.ToString(),
                    Text = e.Name,
                    Selected = this.DepartmentId == e.DepartmentId
                }));
            }
        }

        public void SetAdminUser(EserviceAdminUser eserviceAdminUser)
        {
            if (eserviceAdminUser != null)
            {
                this.EserviceAdminId = eserviceAdminUser.EserviceAdminUserId;
                this.EserviceAdminUsername = eserviceAdminUser.Username;
                this.EserviceAdminName = eserviceAdminUser.Name;
            }
        }

        public void SetDeliveryAdminstrations(DepartmentInstitutionInfo[] deliveryAdministrations)
        {
            if (deliveryAdministrations != null && deliveryAdministrations.Length > 0)
            {
                this.DeliveryAdministrationList = new List<SelectListItem>();
                this.DeliveryAdministrationList.Add(new SelectListItem() { Value = "", Text = "--Не е избрано--" });
                this.DeliveryAdministrationList.AddRange(
                    deliveryAdministrations.Select(o => new SelectListItem()
                    {
                        Value =
                        o.ElectronicSubjectId.ToString(),
                        Text = o.ElectronicSubjectName,
                        Selected = o.ElectronicSubjectId.ToString() == this.DeliveryAdministrationGuid

                    }));


             this.DeliveryAdministrationUIList = deliveryAdministrations
                    .AsEnumerable()
                    .Select(o => new KeyValuePair<string, string>
                    (
                        o.ElectronicSubjectId.ToString(),
                        o.UniqueSubjectIdentifier)
                    )
                    .ToList();
            }
            else
            {
                this.DeliveryAdministrationList = new List<SelectListItem>();
            }
        }
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (this.AggregateToParent == BoolNom.Yes && this.ParentId == null)
            {
                yield return new ValidationResult("АИС клиента не може да агрегира към по-високо ниво и да не е избран първостепенен разпоредител.");
            }

            //if (this.BudgetCode != null)
            //{
            //    if (!int.TryParse(this.BudgetCode, out int number))
            //    {
            //        yield return new ValidationResult("Бюджетният код трябва да е число.");
            //    }
            //}
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

    public enum DistributionType
    {
        [Description("С банков превод")]
        WithBankTransfer = 1,
    }

    public enum ActiveStatus
    {
        [Description("Активен")]
        Activated = 1,

        [Description("Неактивен")]
        Deactivated = 2,
    }
}