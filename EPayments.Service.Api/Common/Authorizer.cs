using EPayments.Data.Repositories.Interfaces;
using EPayments.Service.Api.Common.CustomExceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Common
{
    public class Authorizer
    {
        private IApiRepository apiRepository { get; set; }
        private string clientId { get; set; }

        public Authorizer(IApiRepository apiRepository, string clientId)
        {
            this.apiRepository = apiRepository;
            this.clientId = clientId;
        }

        public void PermitAccessToRequest(string paymentRequestIdentifier)
        {
            bool isAuthorized = apiRepository.IsClientAuthorizedToAccessRequests(this.clientId, new List<string> { paymentRequestIdentifier });
            if (!isAuthorized)
            {
                throw new CustomServiceException("Not authorized to access payment request.");
            }
        }

        public void PermitAccessToRequests(List<string> paymentRequestIdentifiers, bool authorizeIfClientIdIsPaymentInitiator = false)
        {
            bool isAuthorized = apiRepository.IsClientAuthorizedToAccessRequests(this.clientId, paymentRequestIdentifiers, authorizeIfClientIdIsPaymentInitiator);
            if (!isAuthorized)
            {
                throw new CustomServiceException("Not authorized to access payment request.");
            }
        }
    }
}
