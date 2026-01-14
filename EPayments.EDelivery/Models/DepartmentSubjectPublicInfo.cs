using System;
using System.Runtime.Serialization;
using ProdInstitutionPublicInfo = EPayments.EDelivery.EDeliveryProductionClient.DcSubjectPublicInfo;
using TestInstitutionPublicInfo = EPayments.EDelivery.EDeliveryClient.DcSubjectPublicInfo;

namespace EPayments.EDelivery.Models
{
    public class DepartmentSubjectPublicInfo
    {
        public ExtensionDataObject ExtensionData { get; set; }

        public Guid ElectronicSubjectId { get; set; }

        public string ElectronicSubjectName { get; set; }

        public string Email { get; set; }

        public bool IsActivated { get; set; }

        public string PhoneNumber { get; set; }

        public DepartmentProfileTypeEnum ProfileType { get; set; }

        internal static DepartmentSubjectPublicInfo Map(TestInstitutionPublicInfo dcInstitutionInfo)
        {
            if (dcInstitutionInfo == null)
            {
                return null;
            }
            
            return new DepartmentSubjectPublicInfo()
            {
                ExtensionData = dcInstitutionInfo.ExtensionData,
                ElectronicSubjectName = dcInstitutionInfo.ElectronicSubjectName,
                Email = dcInstitutionInfo.Email,
                IsActivated = dcInstitutionInfo.IsActivated,
                PhoneNumber = dcInstitutionInfo.PhoneNumber,
                ProfileType = (DepartmentProfileTypeEnum)((int)dcInstitutionInfo.ProfileType)
            };
        }

        internal static DepartmentSubjectPublicInfo Map(ProdInstitutionPublicInfo dcInstitutionInfo)
        {
            if (dcInstitutionInfo == null)
            {
                return null;
            }

            return new DepartmentSubjectPublicInfo()
            {
                ExtensionData = dcInstitutionInfo.ExtensionData,
                ElectronicSubjectName = dcInstitutionInfo.ElectronicSubjectName,
                Email = dcInstitutionInfo.Email,
                IsActivated = dcInstitutionInfo.IsActivated,
                PhoneNumber = dcInstitutionInfo.PhoneNumber,
                ProfileType = (DepartmentProfileTypeEnum)((int)dcInstitutionInfo.ProfileType)
            };
        }
    }
}
