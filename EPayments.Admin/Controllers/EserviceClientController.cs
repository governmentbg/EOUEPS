using EPayments.Admin.Auth;
using EPayments.Admin.Common;
using EPayments.Admin.DataObjects;
using EPayments.Admin.Models.EserviceClient;
using EPayments.Admin.Models.Shared;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.DataObjects;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web.Helpers;
using System.Web.Mvc;
using OT = EPayments.Admin.Models.ObligationType;
using EPayments.EDelivery.Manager;
using log4net;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.Modify)]
    public partial class EserviceClientController : BaseController
    {
        private IWebRepository webRepository;
        private IUnitOfWork unitOfWork;
        private IDeliveryRegisterManager deliveryManager;
        private static readonly ILog Logger = LogManager.GetLogger(nameof(EserviceClientController));

        public EserviceClientController(IWebRepository webRepository, IUnitOfWork unitOfWork, IDeliveryRegisterManager deliveryManager)
        {
            this.webRepository = webRepository;
            this.unitOfWork = unitOfWork;
            this.deliveryManager = deliveryManager;
        }

        [HttpGet]
        public virtual ActionResult List(ListSearchDO searchDO)
        {
            ListVM model = new ListVM();

            var currentAdminUser = this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .Single(e => e.InternalAdminUserId == this.CurrentUser.UserId);

            model.IsSuperadmin = currentAdminUser.IsSuperadmin;

            model.EserviceRecordsPagingOptions = new PagingVM();

            model.SearchDO = searchDO;

            model.EserviceRecordsPagingOptions.CurrentPageIndex = searchDO.Page;
            model.EserviceRecordsPagingOptions.ControllerName = MVC.EserviceClient.Name;
            model.EserviceRecordsPagingOptions.ActionName = MVC.EserviceClient.ActionNames.List;
            model.EserviceRecordsPagingOptions.PageIndexParameterName = "page";
            model.EserviceRecordsPagingOptions.RouteValues = searchDO.ToRouteValues();

            model.EserviceRecordsPagingOptions.TotalItemCount = this.webRepository.CountEserviceRecords(
                searchDO.AisName,
                searchDO.DepartmentName,
                searchDO.VposClientId,
                searchDO.IsActiveId);

            model.EserviceRecords = this.webRepository.GetEserviceRecords(
                searchDO.AisName,
                searchDO.DepartmentName,
                searchDO.VposClientId,
                searchDO.IsActiveId,
                searchDO.SortBy,
                searchDO.SortDesc,
                searchDO.Page,
                10);

            model.BoricaWarnings = GetBoricaWarnings();

            return View(model);
        }

        private Dictionary<int, string> GetBoricaWarnings()
        {
            var boricaWarnings = new Dictionary<int, string>();

            List<EserviceClient> boricaClientsWithWarning = this.webRepository.GetBoricaEserviceClientWarnings();

            foreach (var boricaClient in boricaClientsWithWarning)
            {
                string msg;

                if (DateTime.Now > boricaClient.BoricaVposRequestSignCertValidTo.Value)
                {
                    msg = $"БОРИКА сертификатът на АИС клиент „{boricaClient.AisName}“ е изтекъл на {Formatter.DateToBgFormat(boricaClient.BoricaVposRequestSignCertValidTo.Value)}.";
                }
                else
                {
                    var days = (boricaClient.BoricaVposRequestSignCertValidTo.Value - DateTime.Now).Days;

                    if (days > 1)
                    {
                        msg = $"БОРИКА сертификатът на АИС клиент „{boricaClient.AisName}“ изтича след {days} дни на {Formatter.DateToBgFormat(boricaClient.BoricaVposRequestSignCertValidTo.Value)}.";
                    }
                    else if (days == 1)
                    {
                        msg = $"БОРИКА сертификатът на АИС клиент „{boricaClient.AisName}“ изтича след 1 ден на {Formatter.DateToBgFormat(boricaClient.BoricaVposRequestSignCertValidTo.Value)}.";
                    }
                    else if (days == 0)
                    {
                        msg = $"БОРИКА сертификатът на АИС клиент „{boricaClient.AisName}“ изтича днес на {Formatter.DateToBgFormat(boricaClient.BoricaVposRequestSignCertValidTo.Value)}.";
                    }
                    else
                    {
                        msg = $"БОРИКА сертификатът на АИС клиент „{boricaClient.AisName}“ е изтекъл на {Formatter.DateToBgFormat(boricaClient.BoricaVposRequestSignCertValidTo.Value)}.";
                    }
                }

                boricaWarnings.Add(boricaClient.EserviceClientId, msg);
            }

            return boricaWarnings;
        }

        [HttpPost]
        public virtual ActionResult ListSearch(ListSearchDO searchDO)
        {
            searchDO.Page = 1;

            return RedirectToListAction(searchDO);
        }

        [HttpGet]
        public virtual ActionResult ListSort(ListSearchDO searchDO)
        {
            return RedirectToListAction(searchDO);
        }

        [HttpGet]
        public virtual ActionResult View(int id)
        {
            try
            {
                EServiceClientVM model = this.CreateModel(id, true, FormMode.View);
                return View(model);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                throw;
            }
        }

        [HttpGet]
        public virtual ActionResult Create()
        {
            try
            {
                EServiceClientVM model = this.CreateModel(null, false, FormMode.Create);
                return View(model);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                throw;
            }
        }

        [HttpPost]
        public virtual ActionResult Create(EServiceClientVM model)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    model = AddModelParameters(model, false, FormMode.Create);
                    return View(model);
                }

                if (this.unitOfWork.DbContext.Set<EserviceClient>().Any(ec => ec.AisName == model.AisName))
                {
                    this.ModelState.AddModelError(string.Empty, "Съществува друг АИС клиент със същото име.");

                    model = AddModelParameters(model, false, FormMode.Create);
                    return View(model);
                }

                var eserviceClient = new EserviceClient();

                eserviceClient.Alias = Guid.NewGuid().ToString();
                eserviceClient.AisName = model.AisName;
                eserviceClient.ServiceName = model.AisName;
                eserviceClient.AccountBank = model.AccountBank;
                eserviceClient.AccountBIC = model.AccountBIC.ToUpper();
                eserviceClient.AccountIBAN = model.AccountIBAN.ToUpper();
                eserviceClient.IsEpayVposEnabled = model.IsEpayVposEnabled;
                eserviceClient.IsBoricaVposEnabled = model.IsBoricaVposEnabled;
                eserviceClient.IsActive = model.IsActiveBoolNom == ActiveStatus.Activated;
                // eserviceClient.ObligationTypeId = model.ObligationTypeId;
                eserviceClient.DeliveryAdminstrationId = model.DeliveryAdminstrationId;
                eserviceClient.DeliveryAdministrationGuid = model.DeliveryAdministrationGuid;

                eserviceClient.DepartmentId = model.DepartmentId;
                eserviceClient.AggregateToParent = model.AggregateToParent == BoolNom.Yes;
                eserviceClient.DistributionTypeId = (int)model.DistributionType;
                eserviceClient.ParentId = model.ParentId;
                eserviceClient.Gid = Guid.NewGuid();
                eserviceClient.IsAuthPassAuthorized = false;
                eserviceClient.ClientId = GenerateClientId(eserviceClient.Gid);
                eserviceClient.SecretKey = GenerateSecretKey();
                //eserviceClient.BudgetCode = model.BudgetCode;

                this.unitOfWork.DbContext.Set<EserviceClient>().Add(eserviceClient);
                this.unitOfWork.Save();

                TempData[TempDataKeys.Message] = "АИС клиента беше успешно добавен.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
                throw;
            }
        }

        [HttpGet]
        public virtual ActionResult Edit(int id)
        {
            var model = this.CreateModel(id, true, FormMode.Edit);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult Edit(int id, EServiceClientVM model)
        {
            if (!ModelState.IsValid)
            {
                model = AddModelParameters(model, true, FormMode.Edit);
                return View(model);
            }

            if (this.unitOfWork.DbContext.Set<EserviceClient>().Any(ec => ec.AisName == model.AisName && ec.EserviceClientId != model.EserviceClientId))
            {
                this.ModelState.AddModelError(string.Empty, "Съществува друг АИС клиент със същото име.");

                model = AddModelParameters(model, true, FormMode.Edit);
                return View(model);
            }

            EserviceClient eserviceClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Include(ec => ec.Children)
                .SingleOrDefault(ec => ec.EserviceClientId == model.EserviceClientId);

            if (eserviceClient.DepartmentId != model.DepartmentId)
            {
                if (eserviceClient.Children.Count > 0)
                {
                    model = AddModelParameters(model, true, FormMode.Edit);

                    this.ModelState.AddModelError(string.Empty, "Не можете да преместите АИС клиента в друга административна структура, защото той е разпоредител.");

                    return View(model);
                }
            }

            if (model.IsActiveBoolNom == ActiveStatus.Deactivated && eserviceClient.Children.Count > 0)
            {
                model = AddModelParameters(model, true, FormMode.Edit);

                this.ModelState.AddModelError(string.Empty, "Не можете да деактивирате АИС клиент, който е разпоредител.");

                return View(model);
            }

            if (eserviceClient.ParentId != model.ParentId && eserviceClient.Children.Count > 0)
            {
                model = AddModelParameters(model, true, FormMode.Edit);

                this.ModelState.AddModelError(string.Empty, "Не можете да промените разпоредителя на АИС клиента, защото той е разпоредител.");

                return View(model);
            }

            if (model.ParentId != null)
            {
                EserviceClient targetClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                    .Include(ec => ec.Children)
                    .SingleOrDefault(ec => ec.EserviceClientId == (int)model.ParentId);

                if (targetClient == null)
                {
                    model = AddModelParameters(model, true, FormMode.Edit);

                    this.ModelState.AddModelError(string.Empty, "Разпоредителя не е намерен.");

                    return View(model);
                }

                if (targetClient.DepartmentId != eserviceClient.DepartmentId)
                {
                    model = AddModelParameters(model, true, FormMode.Edit);

                    this.ModelState.AddModelError(string.Empty, "Разпоредителя не е в същата административна структура.");

                    return View(model);
                }

                if (targetClient.IsActive == false)
                {
                    model = AddModelParameters(model, true, FormMode.Edit);

                    this.ModelState.AddModelError(string.Empty, "Не можете да добавите за разпоредител, неактивен АИС клиент.");

                    return View(model);
                }
            }
            var deliveryAdmins = this.deliveryManager.GetAdministration();
            model.SetDeliveryAdminstrations(deliveryAdmins);

            eserviceClient.AisName = model.AisName;
            eserviceClient.AccountBank = model.AccountBank;
            eserviceClient.AccountBIC = model.AccountBIC.ToUpper();
            eserviceClient.AccountIBAN = model.AccountIBAN.ToUpper();
            eserviceClient.IsEpayVposEnabled = model.IsEpayVposEnabled;
            eserviceClient.IsBoricaVposEnabled = model.IsBoricaVposEnabled;
            eserviceClient.IsActive = model.IsActiveBoolNom == ActiveStatus.Activated;
            //eserviceClient.ObligationTypeId = model.ObligationTypeId;
            eserviceClient.DepartmentId = model.DepartmentId;
            eserviceClient.AggregateToParent = model.AggregateToParent == BoolNom.Yes;
            eserviceClient.DistributionTypeId = (int)model.DistributionType;
            eserviceClient.ParentId = model.ParentId;
            eserviceClient.DeliveryAdministrationGuid = model.DeliveryAdministrationGuid;
            //eserviceClient.BudgetCode = model.BudgetCode;
            eserviceClient.DeliveryAdminstrationId = model.DeliveryAdministrationGuid != null ?
                    model.DeliveryAdministrationUIList.FirstOrDefault(x => x.Key == model.DeliveryAdministrationGuid).Value : 
                    null;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за АИС клиента бяха успешно редактирани.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpPost]
        public virtual ActionResult Delete(int id)
        {
            try
            {
                var eserviceClient = this.webRepository.GetEserviceClient(id);

                if (eserviceClient != null)
                {
                    bool hasChildren = this.unitOfWork.DbContext.Set<EserviceClient>()
                        .Any(ec => ec.ParentId == eserviceClient.EserviceClientId);

                    if (hasChildren)
                    {
                        TempData[TempDataKeys.ErrorMessage] = "Не можете да изтриете АИС клиент, който е разпоредител.";

                        return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
                    }

                    this.unitOfWork.DbContext.Set<EserviceClient>().Remove(eserviceClient);
                    this.unitOfWork.Save();

                    TempData[TempDataKeys.Message] = "АИС клиента беше успешно изтрит.";
                }
                else
                {
                    TempData[TempDataKeys.ErrorMessage] = "АИС клиента не беше намерен.";
                }
            }
            catch
            {
                TempData[TempDataKeys.ErrorMessage] = "АИС клиента не може да бъде изтрит.";
            }

            return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
        }

        [HttpGet]
        public virtual ActionResult AddVpos(int id, Vpos? type)
        {
            if (!type.HasValue)
            {
                TempData[TempDataKeys.ErrorMessage] = "Моля изберте типа на ВПОС терминала, който искате да добавите.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            switch (type.Value)
            {
                case Vpos.Dsk:
                case Vpos.DskEcomm:
                    return RedirectToAction(MVC.EserviceClient.ActionNames.AddVposDsk, MVC.EserviceClient.Name, new { id = id });
                case Vpos.Borica:
                    return RedirectToAction(MVC.EserviceClient.ActionNames.AddVposBorica, MVC.EserviceClient.Name, new { id = id });
                case Vpos.FiBank:
                    return RedirectToAction(MVC.EserviceClient.ActionNames.AddVposFibank, MVC.EserviceClient.Name, new { id = id });
                default:
                    throw new ArgumentException();
            }
        }

        //Vpos Dsk

        [HttpGet]
        public virtual ActionResult AddVposDsk(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposDskVM model = new VposDskVM(FormMode.Create, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult AddVposDsk(int id, VposDskVM model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId.HasValue)
            {
                TempData[TempDataKeys.ErrorMessage] = "Вече има добавен ВПОС терминал за този АИС клиент.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            eserviceClient.VposClientId = (int)Vpos.DskEcomm;
            eserviceClient.DskVposMerchantId = model.DskVposMerchantId;
            eserviceClient.DskVposMerchantPassword = model.DskVposMerchantPassword;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на ДСК бяха успешно добавени.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpGet]
        public virtual ActionResult ViewVposDsk(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposDskVM model = new VposDskVM(FormMode.View, eserviceClient);

            model.TestVposRequestDO = CreateTestVposAuthRequestDO(eserviceClient.ClientId, eserviceClient.SecretKey);
            model.TestVposPostUrl = GetTestVposPostUrl();

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult EditVposDsk(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposDskVM model = new VposDskVM(FormMode.Edit, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult EditVposDsk(int id, VposDskVM model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.DskEcomm && eserviceClient.VposClientId != (int)Vpos.Dsk)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            eserviceClient.DskVposMerchantId = model.DskVposMerchantId;
            eserviceClient.DskVposMerchantPassword = model.DskVposMerchantPassword;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на ДСК бяха успешно редактирани.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.ViewVposDsk, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpPost]
        public virtual ActionResult DeleteVposDsk(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.DskEcomm && eserviceClient.VposClientId != (int)Vpos.Dsk)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.ViewVposDsk, MVC.EserviceClient.Name, new { id = id });
            }

            eserviceClient.VposClientId = null;
            eserviceClient.DskVposMerchantId = null;
            eserviceClient.DskVposMerchantPassword = null;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на ДСК бяха изтрити.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        //Vpos Borica

        [HttpGet]
        public virtual ActionResult AddVposBorica(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposBoricaVM model = new VposBoricaVM(FormMode.Create, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult AddVposBorica(int id, VposBoricaVM model)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId.HasValue)
            {
                TempData[TempDataKeys.ErrorMessage] = "Вече има добавен ВПОС терминал за този АИС клиент.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            string certificateFilePath = null;
            string certificateFilePassword = model.BoricaVposRequestSignCertPassword ?? string.Empty;
            DateTime? certificateFileValidTo = null;

            if (model.BoricaVposRequestSignCertFile == null)
            {
                ModelState.AddModelError(string.Empty, "Полето „P12 сертификат за подписване“ е задължително.");
            }
            else if (!model.BoricaVposRequestSignCertFile.FileName.TrimEnd().ToLower().EndsWith(".p12"))
            {
                ModelState.AddModelError(string.Empty, "Невалиден формат на файла. Сертификатът трябва да бъде с разширение „.p12“.");
            }
            else if (model.BoricaVposRequestSignCertFile.ContentLength == 0)
            {
                ModelState.AddModelError(string.Empty, "P12 сертификат файла е празен.");
            }
            else
            {
                certificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_BoricaCertificateFolder, model.BoricaVposRequestSignCertFile.FileName);

                if (System.IO.File.Exists(certificateFilePath))
                {
                    ModelState.AddModelError(string.Empty, "Файл с такова име вече съществува.");
                }
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            model.BoricaVposRequestSignCertFile.SaveAs(certificateFilePath);

            //test certificate password
            try
            {
                X509Certificate2Collection collection = new X509Certificate2Collection();
                collection.Import(certificateFilePath, certificateFilePassword,
                    AppSettings.EPaymentsCommon_UseMachineKeySet == false ?
                    X509KeyStorageFlags.PersistKeySet :
                    X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.MachineKeySet);

                certificateFileValidTo = collection[0].NotAfter;
            }
            catch
            {
                System.IO.File.Delete(certificateFilePath);

                ModelState.AddModelError(string.Empty, "Ключа за P12 сертификата е невалиден.");
                return View(model);
            }

            eserviceClient.VposClientId = (int)Vpos.Borica;

            eserviceClient.BoricaVposRequestSignCertFileName = model.BoricaVposRequestSignCertFile.FileName;
            eserviceClient.BoricaVposRequestSignCertPassword = certificateFilePassword;
            eserviceClient.BoricaVposRequestSignCertValidTo = certificateFileValidTo;

            eserviceClient.BoricaVposTerminalId = model.BoricaVposTerminalId;
            eserviceClient.BoricaVposMerchantId = model.BoricaVposMerchantId;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на БОРИКА бяха успешно добавени.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpGet]
        public virtual ActionResult ViewVposBorica(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposBoricaVM model = new VposBoricaVM(FormMode.View, eserviceClient);

            model.TestVposRequestDO = CreateTestVposAuthRequestDO(eserviceClient.ClientId, eserviceClient.SecretKey);
            model.TestVposPostUrl = GetTestVposPostUrl();

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult EditVposBorica(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposBoricaVM model = new VposBoricaVM(FormMode.Edit, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult EditVposBorica(int id, VposBoricaVM model)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.Borica)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            string certificateFilePassword = model.BoricaVposRequestSignCertPassword ?? string.Empty;

            if (model.BoricaVposRequestSignCertFile == null)
            {
                if (certificateFilePassword != eserviceClient.BoricaVposRequestSignCertPassword)
                {
                    ModelState.AddModelError(string.Empty, "Ключа за сертификата може да бъде сменян само, когато се прикачва сертификата.");
                }
            }
            else
            {
                string certificateFilePath = null;
                DateTime? certificateFileValidTo = null;

                if (!model.BoricaVposRequestSignCertFile.FileName.TrimEnd().ToLower().EndsWith(".p12"))
                {
                    ModelState.AddModelError(string.Empty, "Невалиден формат на файла. Сертификатът трябва да бъде с разширение „.p12“.");
                }
                else if (model.BoricaVposRequestSignCertFile.ContentLength == 0)
                {
                    ModelState.AddModelError(string.Empty, "Конфигурационния файл е празен.");
                }
                else
                {
                    certificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_BoricaCertificateFolder, model.BoricaVposRequestSignCertFile.FileName);

                    if (model.BoricaVposRequestSignCertFile.FileName.ToLower().Trim() != eserviceClient.BoricaVposRequestSignCertFileName?.ToLower().Trim()
                        && System.IO.File.Exists(certificateFilePath))
                    {
                        ModelState.AddModelError(string.Empty, "Файл с такова име вече съществува.");
                    }
                }

                if (!ModelState.IsValid)
                {
                    return View(model);
                }

                string tempCertificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_BoricaCertificateFolder, $"tempcert_{Guid.NewGuid()}.p12");
                model.BoricaVposRequestSignCertFile.SaveAs(tempCertificateFilePath);

                //test certificate password
                try
                {
                    X509Certificate2Collection collection = new X509Certificate2Collection();
                    collection.Import(tempCertificateFilePath,
                        certificateFilePassword,
                        AppSettings.EPaymentsCommon_UseMachineKeySet == false ?
                    X509KeyStorageFlags.PersistKeySet :
                    X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.MachineKeySet);

                    certificateFileValidTo = collection[0].NotAfter;
                }
                catch
                {
                    System.IO.File.Delete(tempCertificateFilePath);

                    ModelState.AddModelError(string.Empty, "Ключа за P12 сертификата е невалиден.");
                    return View(model);
                }

                //delete existing
                string existingCertificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_BoricaCertificateFolder, eserviceClient.BoricaVposRequestSignCertFileName);
                if (System.IO.File.Exists(existingCertificateFilePath))
                {
                    System.IO.File.Delete(existingCertificateFilePath);
                }

                System.IO.File.Move(tempCertificateFilePath, certificateFilePath);

                eserviceClient.BoricaVposRequestSignCertFileName = model.BoricaVposRequestSignCertFile.FileName;
                eserviceClient.BoricaVposRequestSignCertPassword = certificateFilePassword;
                eserviceClient.BoricaVposRequestSignCertValidTo = certificateFileValidTo;
                eserviceClient.BoricaVposRequestSignCertExpHideAdminMsg = null;
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            eserviceClient.BoricaVposTerminalId = model.BoricaVposTerminalId;
            eserviceClient.BoricaVposMerchantId = model.BoricaVposMerchantId;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на БОРИКА бяха успешно редактирани.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.ViewVposBorica, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpPost]
        public virtual ActionResult DeleteVposBorica(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.Borica)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.ViewVposBorica, MVC.EserviceClient.Name, new { id = id });
            }

            string existingCertificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_BoricaCertificateFolder, eserviceClient.BoricaVposRequestSignCertFileName);
            if (System.IO.File.Exists(existingCertificateFilePath))
            {
                System.IO.File.Delete(existingCertificateFilePath);
            }

            eserviceClient.VposClientId = null;

            eserviceClient.BoricaVposRequestSignCertFileName = null;
            eserviceClient.BoricaVposRequestSignCertPassword = null;
            eserviceClient.BoricaVposRequestSignCertValidTo = null;

            eserviceClient.BoricaVposTerminalId = null;
            eserviceClient.BoricaVposMerchantId = null;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на БОРИКА бяха изтрити.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpGet]
        public virtual ActionResult HideVposBoricaWarning(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.Borica)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
            }

            eserviceClient.BoricaVposRequestSignCertExpHideAdminMsg = true;

            this.unitOfWork.Save();

            return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name);
        }

        //Vpos Fibank

        [HttpGet]
        public virtual ActionResult AddVposFibank(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposFibankVM model = new VposFibankVM(FormMode.Create, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult AddVposFibank(int id, VposFibankVM model)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId.HasValue)
            {
                TempData[TempDataKeys.ErrorMessage] = "Вече има добавен ВПОС терминал за този АИС клиент.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            string certificateFilePath = null;
            string filePrefix = "FiBankCert_" + Guid.NewGuid().ToString().Substring(0, 6) + "_";

            if (model.FiBankVposAccessKeystoreFile == null)
            {
                ModelState.AddModelError(string.Empty, "Полето „Конфигурационен файл“ е задължително.");
            }
            else if (model.FiBankVposAccessKeystoreFile.ContentLength == 0)
            {
                ModelState.AddModelError(string.Empty, "Конфигурационния файл е празен.");
            }
            else
            {
                certificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_FiBankCertificateFolder, filePrefix + model.FiBankVposAccessKeystoreFile.FileName);

                if (System.IO.File.Exists(certificateFilePath))
                {
                    ModelState.AddModelError(string.Empty, "Файл с такова име вече съществува.");
                }
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            model.FiBankVposAccessKeystoreFile.SaveAs(certificateFilePath);

            eserviceClient.VposClientId = (int)Vpos.FiBank;
            eserviceClient.FiBankVposAccessKeystorePassword = model.FiBankVposAccessKeystorePassword;
            eserviceClient.FiBankVposAccessKeystoreFilename = filePrefix + model.FiBankVposAccessKeystoreFile.FileName;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на ПИБ бяха успешно добавени.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpGet]
        public virtual ActionResult ViewVposFibank(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposFibankVM model = new VposFibankVM(FormMode.View, eserviceClient);

            model.TestVposRequestDO = CreateTestVposAuthRequestDO(eserviceClient.ClientId, eserviceClient.SecretKey);
            model.TestVposPostUrl = GetTestVposPostUrl();

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult EditVposFibank(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            VposFibankVM model = new VposFibankVM(FormMode.Edit, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult EditVposFibank(int id, VposFibankVM model)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.FiBank)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            if (model.FiBankVposAccessKeystoreFile != null)
            {
                string certificateFilePath = null;
                string filePrefix = "FiBankCert_" + Guid.NewGuid().ToString().Substring(0, 6) + "_";

                if (model.FiBankVposAccessKeystoreFile.ContentLength == 0)
                {
                    ModelState.AddModelError(string.Empty, "Конфигурационния файл е празен.");
                }
                else
                {
                    certificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_FiBankCertificateFolder, filePrefix + model.FiBankVposAccessKeystoreFile.FileName);

                    if (model.FiBankVposAccessKeystoreFile.FileName.ToLower().Trim() != eserviceClient.FiBankVposAccessKeystoreFilename?.ToLower().Trim()
                        && System.IO.File.Exists(certificateFilePath))
                    {
                        ModelState.AddModelError(string.Empty, "Файл с такова име вече съществува.");
                    }
                }

                if (!ModelState.IsValid)
                {
                    return View(model);
                }

                //delete existing
                string existingCertificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_FiBankCertificateFolder, eserviceClient.FiBankVposAccessKeystoreFilename);
                if (System.IO.File.Exists(existingCertificateFilePath))
                {
                    try
                    {
                        System.IO.File.Delete(existingCertificateFilePath);
                    }
                    catch { }

                }

                model.FiBankVposAccessKeystoreFile.SaveAs(certificateFilePath);

                eserviceClient.FiBankVposAccessKeystoreFilename = filePrefix + model.FiBankVposAccessKeystoreFile.FileName;
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            eserviceClient.FiBankVposAccessKeystorePassword = model.FiBankVposAccessKeystorePassword;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на ПИБ бяха успешно редактирани.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.ViewVposFibank, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpPost]
        public virtual ActionResult DeleteVposFibank(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceClient.VposClientId != (int)Vpos.FiBank)
            {
                TempData[TempDataKeys.ErrorMessage] = "Възникна грешка.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.ViewVposFibank, MVC.EserviceClient.Name, new { id = id });
            }

            string existingCertificateFilePath = Path.Combine(AppSettings.EPaymentsAdmin_FiBankCertificateFolder, eserviceClient.FiBankVposAccessKeystoreFilename);
            if (System.IO.File.Exists(existingCertificateFilePath))
            {
                try
                {
                    System.IO.File.Delete(existingCertificateFilePath);
                }
                catch { }
            }

            eserviceClient.VposClientId = null;
            eserviceClient.FiBankVposAccessKeystoreFilename = null;
            eserviceClient.FiBankVposAccessKeystorePassword = null;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за ВПОС терминала на ПИБ бяха изтрити.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        //E-Forms

        [HttpGet]
        public virtual ActionResult AddEformsUser(int id)
        {
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            EformsUserVM model = new EformsUserVM(FormMode.Create, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult AddEformsUser(int id, EformsUserVM model)
        {
            if (ModelState.IsValid)
            {
                if (this.webRepository.EserviceAdminUserExists(model.Username))
                {
                    ModelState.AddModelError(string.Empty, "Вече съществува потребител със същото потребителско име.");
                }
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            EserviceAdminUser eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);

            if (eserviceAdminUser != null)
            {
                TempData[TempDataKeys.ErrorMessage] = "Вече има добавен системен потребител за Е-Форми за този АИС клиент.";

                return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
            }

            var eserviceClient = this.webRepository.GetEserviceClient(id);

            string salt = Crypto.GenerateSalt();
            string hash = Crypto.HashPassword(model.Password + salt);

            eserviceAdminUser = new EserviceAdminUser();
            eserviceAdminUser.ReferringEserviceClientId = id;
            eserviceAdminUser.Username = model.Username.Trim();
            eserviceAdminUser.Name = model.Name.Trim();
            eserviceAdminUser.PasswordHash = hash;
            eserviceAdminUser.PasswordSalt = salt;
            eserviceAdminUser.IsActive = true;

            this.unitOfWork.DbContext.Set<EserviceAdminUser>().Add(eserviceAdminUser);
            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Системният потребител за Е-Форми беше успешно добавен.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpGet]
        public virtual ActionResult ViewEformsUser(int id, int eserviceAdminUserId)
        {
            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceAdminUser.EserviceAdminUserId != eserviceAdminUserId)
                throw new ArgumentException();

            EformsUserVM model = new EformsUserVM(FormMode.View, eserviceAdminUser, eserviceClient);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult EditEformsUser(int id, int eserviceAdminUserId)
        {
            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceAdminUser.EserviceAdminUserId != eserviceAdminUserId)
                throw new ArgumentException();

            EformsUserVM model = new EformsUserVM(FormMode.Edit, eserviceAdminUser, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult EditEformsUser(int id, int eserviceAdminUserId, EformsUserVM model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);

            if (eserviceAdminUser.EserviceAdminUserId != eserviceAdminUserId)
                throw new ArgumentException();

            eserviceAdminUser.Username = model.Username.Trim();
            eserviceAdminUser.Name = model.Name.Trim();

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Данните за системения потребител за Е-Форми бяха успешно редактирани.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.ViewEformsUser, MVC.EserviceClient.Name, new { id = id, eserviceAdminUserId = eserviceAdminUserId });
        }

        [HttpGet]
        public virtual ActionResult EditEformsUserPassword(int id, int eserviceAdminUserId)
        {
            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);
            var eserviceClient = this.webRepository.GetEserviceClient(id);

            if (eserviceAdminUser.EserviceAdminUserId != eserviceAdminUserId)
                throw new ArgumentException();

            EformsUserPasswordVM model = new EformsUserPasswordVM(eserviceAdminUser, eserviceClient);

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult EditEformsUserPassword(int id, int eserviceAdminUserId, EformsUserPasswordVM model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);

            if (eserviceAdminUser.EserviceAdminUserId != eserviceAdminUserId)
                throw new ArgumentException();

            string salt = Crypto.GenerateSalt();
            string hash = Crypto.HashPassword(model.Password + salt);

            eserviceAdminUser.PasswordHash = hash;
            eserviceAdminUser.PasswordSalt = salt;

            this.unitOfWork.Save();

            TempData[TempDataKeys.Message] = "Паролата на системения потребител за Е-Форми беше успешно сменена.";

            return RedirectToAction(MVC.EserviceClient.ActionNames.ViewEformsUser, MVC.EserviceClient.Name, new { id = id, eserviceAdminUserId = eserviceAdminUserId });
        }

        [HttpPost]
        public virtual ActionResult DeleteEformsUser(int id, int eserviceAdminUserId)
        {
            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(id);

            if (eserviceAdminUser.EserviceAdminUserId != eserviceAdminUserId)
                throw new ArgumentException();

            try
            {
                this.unitOfWork.DbContext.Set<EserviceAdminUser>().Remove(eserviceAdminUser);
                this.unitOfWork.Save();

                TempData[TempDataKeys.Message] = "Системният потребител за Е-Форми беше успешно изтрит.";
            }
            catch
            {
                TempData[TempDataKeys.ErrorMessage] = "Системният потребител за Е-Форми не може да бъде изтрит.";
            }

            return RedirectToAction(MVC.EserviceClient.ActionNames.View, MVC.EserviceClient.Name, new { id = id });
        }

        [HttpGet]
        public virtual FileResult DownloadUserManual()
        {
            string localFileName = "UserManual_v.1.0.docx";
            string saveFileName = "UserManual_v.1.0.docx";
            string mimeType = MimeTypeFileExtension.MIME_APPLICATION_MSWORD;

            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", localFileName);

            byte[] data = System.IO.File.ReadAllBytes(filePath);

            return File(data, mimeType, saveFileName);
        }

        [NonAction]
        private ActionResult RedirectToListAction(ListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceClient.ActionNames.List, MVC.EserviceClient.Name,
                new
                {
                    @aisName = searchDO.AisName,
                    @departmentName = searchDO.DepartmentName,
                    @vposClientId = searchDO.VposClientId,
                    @isActiveId = searchDO.IsActiveId,

                    @page = searchDO.Page,
                    @sortBy = searchDO.SortBy,
                    @sortDesc = searchDO.SortDesc,
                });
        }

        private string GenerateSecretKey()
        {
            CodeGenerator codeGenerator = new CodeGenerator();
            codeGenerator.Minimum = 32;
            codeGenerator.Maximum = 32;
            codeGenerator.ConsecutiveCharacters = true;
            codeGenerator.RepeatCharacters = true;

            return codeGenerator.Generate();
        }

        private string GenerateClientId(Guid guid)
        {
            return $"epayments_ais_client_{guid.ToString().ToLower()}";
        }

        private AuthRequestDO CreateTestVposAuthRequestDO(string clientId, string secretKey)
        {
            var jsonDataBytes = Encoding.UTF8.GetBytes(Guid.NewGuid().ToString());

            var base64Data = Convert.ToBase64String(jsonDataBytes);

            var hmac = CalculateHmac(secretKey, base64Data);

            AuthRequestDO requestDO = new AuthRequestDO();
            requestDO.ClientId = clientId;
            requestDO.Hmac = hmac;
            requestDO.Data = base64Data;

            return requestDO;
        }

        private string GetTestVposPostUrl()
        {
            return Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, "Vpos/TestEserviceClientVpos").ToString();
        }

        private EServiceClientVM AddModelParameters(EServiceClientVM model, bool addClients, FormMode formMode)
        {
            if (model == null)
            {
                throw new ArgumentNullException(nameof(EServiceClientVM));
            }

            var departments = this.unitOfWork.DbContext.Set<Department>().ToList();
            var deliveryAdmins = this.deliveryManager.GetAdministration();
            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>()
                .Where(ot => ot.IsActive)
                .Select(OT.ObligationTypeOptionVM.MapFrom)
                .ToList();

            var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(model.EserviceClientId);

            model.SetDeparments(departments);
            //model.SetObligationTypes(obligationTypes);
            model.SetAdminUser(eserviceAdminUser);
            model.Mode = formMode;

            model.SetDeliveryAdminstrations(deliveryAdmins);
            if (addClients == true)
            {
                var eserviceClients = this.unitOfWork.DbContext.Set<EserviceClient>()
                    .Where(ec => ec.DepartmentId == model.DepartmentId && ec.IsActive)
                    .ToList();

                model.SetServiceEClients(eserviceClients);
            }

            return model;
        }

        private EServiceClientVM CreateModel(int? id, bool addClients, FormMode formMode)
        {
            Logger.Debug("CreateModel() - start.");
            var departments = this.unitOfWork.DbContext.Set<Department>().ToList();
            Logger.Debug("CreateModel() - got Departments.");
            var deliveryAdmins = this.deliveryManager.GetAdministration();
            Logger.Debug("CreateModel() - got Administrations.");
            var obligationTypes = this.unitOfWork.DbContext.Set<ObligationType>()
                .Where(ot => ot.IsActive)
                .Select(OT.ObligationTypeOptionVM.MapFrom)
                .ToList();
            Logger.Debug("CreateModel() - got OblidationTypes.");

            EServiceClientVM model;

            if (id == null)
            {
                Logger.Debug("CreateModel() - creating EServiceClientVM without id.");
                model = new EServiceClientVM(FormMode.Create, null, null, departments, obligationTypes, deliveryAdmins, null);
            }
            else
            {
                Logger.Debug("CreateModel() - creating EServiceClientVM with id.");
                var eserviceClient = this.webRepository.GetEserviceClient((int)id);

                var eserviceAdminUser = this.webRepository.GetEserviceAdminUserByReferringEserviceClientId(eserviceClient.EserviceClientId);
                Logger.Debug("CreateModel() - got eserviceClient by id and eserviceAdminUser by EserviceClientId.");
                
                var clients = addClients == true ? this.unitOfWork.DbContext.Set<EserviceClient>()
                    .Where(ec => ec.DepartmentId == eserviceClient.DepartmentId)
                    .ToList() :
                    null;
                Logger.Debug("CreateModel() - got and filtered EserviceClients with the same department id as the eserviceClient which id was passed.");

                model = new EServiceClientVM(formMode, eserviceClient, eserviceAdminUser, departments, obligationTypes, deliveryAdmins, clients);
            }
            Logger.Debug("CreateModel() - created EServiceClientVM.");

            return model;
        }

        public static string CalculateHmac(string secret, string value)
        {
            var secretBytes = Encoding.UTF8.GetBytes(secret);
            var valueBytes = Encoding.UTF8.GetBytes(value);
            string signature;

            using (var hmac = new HMACSHA256(secretBytes))
            {
                var hash = hmac.ComputeHash(valueBytes);
                signature = Convert.ToBase64String(hash);
            }

            return signature;
        }
    }
}