using System;
using System.Linq.Expressions;
using OType = EPayments.Model.Models.ObligationType;

namespace EPayments.Admin.Models.ObligationType
{
    public class ObligationTypeOptionVM
    {
        public int ObligationTypeId { get; set; }
        
        public string Name { get; set; }

        public static Expression<Func<OType, ObligationTypeOptionVM>> MapFrom { get; } = ot =>
            new ObligationTypeOptionVM()
            {
                ObligationTypeId = ot.ObligationTypeId,
                Name = ot.Name
            };

        public static Expression<Func<ObligationTypeOptionVM, OType>> MapTo { get; } = vm =>
            new OType()
            {
                ObligationTypeId = vm.ObligationTypeId,
                Name = vm.Name,
                IsActive = true
            };
    }
}