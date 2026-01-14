using EPayments.EDelivery.Models;
using EPayments.Model.Models;
using System.Threading.Tasks;

namespace EPayments.EDelivery.Manager
{
    public interface IDeliveryRegisterManager
    {
        DepartmentInstitutionInfo[] GetAdministration();
        Task<DepartmentInstitutionInfo[]> GetAdministrationAsync();
        Task<int> SendMessageAsync(EserviceDeliveryNotification notification, string message);
        Task<int> SendMessagePersonAsync(EserviceDeliveryNotification notification, string message, bool isInstitution = false);
    }
}
