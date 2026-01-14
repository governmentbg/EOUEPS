using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Models;

namespace EPayments.Service.Api.Auth.Attributes
{
    public class EbankingAuthorizeAttribute : HmacAuthorizeAttribute
    {
        protected override HmacClientDO GetHmacClientDO(ISystemRepository systemRepository, string clientId)
        {
            HmacClientDO returnClientDO = null;

            EbankingClient client = systemRepository.GetEbankingClientByClientId(clientId);
            if (client != null)
            {
                returnClientDO = new HmacClientDO();
                returnClientDO.ClientId = clientId;
                returnClientDO.SecretKey = client.SecretKey;
                returnClientDO.IsActive = client.IsActive;
            }

            return returnClientDO;
        }
    }
}
