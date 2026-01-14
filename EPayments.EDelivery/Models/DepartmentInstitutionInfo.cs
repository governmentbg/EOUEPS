using System;
using System.Runtime.Serialization;

namespace EPayments.EDelivery.Models
{
    public class DepartmentInstitutionInfo
    {
        public DepartmentSubjectPublicInfo HeadInstitution { get; set; }

        public string Name { get; set; }

        public DepartmentSubjectPublicInfo[] SubInstitutions { get; set; }

        public ExtensionDataObject ExtensionData { get; set; }

        public Guid ElectronicSubjectId { get; set; }

        public string ElectronicSubjectName { get; set; } 

        public string Email { get; set; } 

        public bool IsActivated { get; set; }

        public string PhoneNumber { get; set; }

        public DepartmentProfileTypeEnum ProfileType { get; set; }

        public DateTime? DateCreated { get; set; } 

        public string UniqueSubjectIdentifier { get; set; } 

        public object[] VerificationInfo { get; set; } 
    }
}
