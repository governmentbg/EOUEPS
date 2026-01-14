using EPayments.Data.ViewObjects.Api;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface IApiRepository : IBaseRepository
    {
        IList<RequestXmlVO> GetRequestXmlsByIndetifiers(List<string> identifiers);

        IList<RequestStatusVO> GetRequestStatusesByIdentifiers(List<string> identifiers);

        List<Tuple<int, RequestInfoVO>> GetRequestInfosByApplicant(UinType applicantUinTypeId, string applicantUin);

        List<Tuple<int, RequestInfoParsedVO>> GetParsedRequestInfosByApplicant(UinType applicantUinTypeId, string applicantUin);

        Tuple<int, RequestInfoVO> GetRequestInfoByAccessCode(string accessCode);

        Tuple<int, RequestInfoParsedVO> GetParsedRequestInfoByAccessCode(string accessCode);

        bool IsValidRequestWithKeyDataExist(string serviceProviderIBAN, string paymentReferenceNumber, DateTime paymentReferenceDate);

        bool IsClientAuthorizedToAccessRequests(string clientId, List<string> paymentRequestIdentifiers, bool authorizeIfClientIdIsPaymentInitiator = false);

        PaymentRequest GetPaymentRequestByAisPaymentId(int eserviceClientId, string aisPaymentId);

        Tuple<int, List<Tuple<PaymentRequestStatus?, List<RequestPaymentInfoParsedVO>>>> GetParsedPaymentRequestInfosByApplicantGrouped(string applicantUin, int pageNumber, int pageSize, PaymentRequestStatus[] paymentStatuses);

        Tuple<int, List<Tuple<int, RequestPaymentInfoParsedVO>>> GetParsedPendingPaymentRequestInfosByApplicant(string applicantUin, int pageNumber, int pageSize);

        Tuple<int, List<Tuple<int, RequestPaymentInfoParsedVO>>> GetParsedPaymentRequestInfosByApplicant(string applicantUin, int pageNumber, int pageSize);

        List<Tuple<int, RequestEikInfoVO>> GetParsedRequestInfoByEik(string eiksnumber);

        List<Tuple<int, RequestRefidInfoVO>> GetParsedRequestInfoByRefid(int refid, string clientId);

        int GetNumberOfPaymentsByMonth(string request_type, string month);

        int GetNumberOfPaymentsFromDateToDate(string request_type, string startDate, string endDate);

        object GetNumberOfVposePayCvposPaidRequestsFromDateToDate(string start_date, string end_date);

        object GetAdministrationsWithTheHighestNumberOfPaymentsFromDateToDate(string start_date, string end_date);

        List<RequestEikInfoVO> GetServiceClientByEik(string eiknumber);

        List<ObligationType> GetAllObligationTypes();
    }
}
