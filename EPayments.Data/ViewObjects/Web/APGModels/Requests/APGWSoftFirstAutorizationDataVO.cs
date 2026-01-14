using System.ComponentModel.DataAnnotations;

namespace EPayments.Data.ViewObjects.Web.APGModels.Requests
{
    public class APGWSoftFirstAutorizationDataVO : APGWSoftRequestDataVO
    {
        [Display(Name = "TRTYPE")]
        public override TrTypeEnum TrType => TrTypeEnum.Authorization;
    }
}
