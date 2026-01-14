using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using System;

namespace EPayments.Data.ViewObjects.Web
{
    [Serializable()]
    public class EserviceRecordVO
    {
        public int EserviceClientId { get; set; }
        public string AisName { get; set; }
        public int DepartmentId { get; set; }
        public string DepartmentName { get; set; }
        public string AccountBank { get; set; }
        public int? VposClientId { get; set; }
        public bool IsActive { get; set; }

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
    }
}
