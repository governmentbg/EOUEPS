using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Client = EPayments.Model.Models.EserviceClient;

namespace EPayments.Admin.Models.EserviceClient
{
    public class EserviceVM
    {
        public int EserviceClientId { get; set; }
        public string AisName { get; set; }
        public int DepartmentId { get; set; }
        public string DepartmentName { get; set; }
        public string AccountBank { get; set; }
        public int? VposClientId { get; set; }
        public bool IsActive { get; set; }
        public int? ParentId { get; set; }

        public ICollection<EserviceVM> Children { get; set; } = new HashSet<EserviceVM>();

        public string GetVposDisplayName()
        {
            if (this.VposClientId.HasValue)
            {
                Vpos vposClientId = (Vpos)this.VposClientId.Value;

                return vposClientId.GetDescription();
            }
            else
            {
                return null;
            }
        }

        public static Expression<Func<Client, EserviceVM>> Map { get; } = (e) =>
           new EserviceVM()
           {
               EserviceClientId = e.EserviceClientId,
               AisName = e.AisName,
               DepartmentId = e.DepartmentId,
               DepartmentName = e.Department.Name,
               AccountBank = e.AccountBank,
               VposClientId = e.VposClientId,
               IsActive = e.IsActive,
               ParentId = e.ParentId
           };
    }
}