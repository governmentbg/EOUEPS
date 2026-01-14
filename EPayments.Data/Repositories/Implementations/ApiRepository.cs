using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using System.Collections.Generic;
using System.Linq;
using EPayments.Model.Models;
using EPayments.Model.Enums;
using EPayments.Data.ViewObjects.Api;
using System;
using System.Data.Entity;

namespace EPayments.Data.Repositories.Implementations
{
    internal class ApiRepository : BaseRepository, IApiRepository
    {
        public ApiRepository(IUnitOfWork unitOfWork)
            : base(unitOfWork)
        {
        }

        public IList<RequestXmlVO> GetRequestXmlsByIndetifiers(List<string> identifiers)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    join prx in this.unitOfWork.DbContext.Set<PaymentRequestXml>() on pr.PaymentRequestXmlId equals prx.PaymentRequestXmlId
                    where identifiers.Contains(pr.PaymentRequestIdentifier)
                    select new RequestXmlVO
                    {
                        Id = pr.PaymentRequestIdentifier,
                        RequestXml = prx.RequestContent,
                        PaymentRequest = pr
                    })
                   .ToList();
        }

        public IList<RequestStatusVO> GetRequestStatusesByIdentifiers(List<string> identifiers)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    where identifiers.Contains(pr.PaymentRequestIdentifier)
                    select new RequestStatusVO
                    {
                        Id = pr.PaymentRequestIdentifier,
                        Status = pr.PaymentRequestStatusId,
                        ChangeTime = pr.PaymentRequestStatusChangeTime
                    })
                    .ToList();
        }

        public List<Tuple<int, RequestInfoVO>> GetRequestInfosByApplicant(UinType applicantUinTypeId, string applicantUin)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    join prx in this.unitOfWork.DbContext.Set<PaymentRequestXml>() on pr.PaymentRequestXmlId equals prx.PaymentRequestXmlId
                    where pr.ApplicantUinTypeId == applicantUinTypeId && pr.ApplicantUin == applicantUin && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending
                    select new
                    {
                        paymentRequestId = pr.PaymentRequestId,
                        requestInfoVO = new RequestInfoVO
                        {
                            Id = pr.PaymentRequestIdentifier,
                            Status = pr.PaymentRequestStatusId,
                            ChangeTime = pr.PaymentRequestStatusChangeTime,
                            RequestXml = prx.RequestContent
                        }
                    })
                    .ToList()
                    .Select(e => new Tuple<int, RequestInfoVO>(e.paymentRequestId, e.requestInfoVO))
                    .ToList();
        }

        public List<Tuple<int, RequestInfoParsedVO>> GetParsedRequestInfosByApplicant(UinType applicantUinTypeId, string applicantUin)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    where pr.ApplicantUinTypeId == applicantUinTypeId && pr.ApplicantUin == applicantUin && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending
                    select new
                    {
                        paymentRequestId = pr.PaymentRequestId,
                        requestInfoParsedVO = new RequestInfoParsedVO
                        {
                            Id = pr.PaymentRequestIdentifier,
                            Status = pr.PaymentRequestStatusId,
                            StatusChangeTime = pr.PaymentRequestStatusChangeTime,
                            ServiceProviderName = pr.ServiceProviderName,
                            ServiceProviderBank = pr.ServiceProviderBank,
                            ServiceProviderBIC = pr.ServiceProviderBIC,
                            ServiceProviderIBAN = pr.ServiceProviderIBAN,
                            Currency = pr.Currency,
                            PaymentTypeCode = pr.PaymentTypeCode,
                            PaymentAmount = pr.PaymentAmount,
                            PaymentReason = pr.PaymentReason,
                            ApplicantUinType = pr.ApplicantUinTypeId,
                            ApplicantUin = pr.ApplicantUin,
                            ApplicantName = pr.ApplicantName,
                            PaymentReferenceType = pr.PaymentReferenceType,
                            PaymentReferenceNumber = pr.PaymentReferenceNumber,
                            PaymentReferenceDate = pr.PaymentReferenceDate,
                            ExpirationDate = pr.ExpirationDate,
                            AdditionalInformation = pr.AdditionalInformation,
                            CreateDate = pr.CreateDate
                        }
                    })
                    .ToList()
                    .Select(e => new Tuple<int, RequestInfoParsedVO>(e.paymentRequestId, e.requestInfoParsedVO))
                    .ToList();
        }

        public Tuple<int, List<Tuple<int, RequestPaymentInfoParsedVO>>> GetParsedPendingPaymentRequestInfosByApplicant(string applicantUin, int pageNumber, int pageSize)
        {
            int page = pageNumber > 0 ? pageNumber : 1;
            int size = pageSize > 0 ? pageSize : 20;

            var allResults = (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                              join es in unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals es.EserviceClientId
                              where (pr.ApplicantUinTypeId == UinType.Egn || pr.ApplicantUinTypeId == UinType.Lnch)
                              && pr.ApplicantUin == applicantUin && pr.PaymentRequestStatusId == PaymentRequestStatus.Pending
                              select new
                              {
                                  paymentRequestId = pr.PaymentRequestId,
                                  requestPaymentInfoParsedVO = new RequestPaymentInfoParsedVO
                                  {
                                      Id = pr.PaymentRequestIdentifier,
                                      EserviceClientAisName = es.AisName,
                                      EserviceClientServiceName = es.ServiceName,
                                      EserviceClientDepartmentId = es.DepartmentId.ToString(),
                                      Status = pr.PaymentRequestStatusId,
                                      StatusChangeTime = pr.PaymentRequestStatusChangeTime,
                                      ServiceProviderName = pr.ServiceProviderName,
                                      ServiceProviderBank = pr.ServiceProviderBank,
                                      ServiceProviderBIC = pr.ServiceProviderBIC,
                                      ServiceProviderIBAN = pr.ServiceProviderIBAN,
                                      Currency = pr.Currency,
                                      PaymentTypeCode = pr.PaymentTypeCode,
                                      PaymentAmount = pr.PaymentAmount,
                                      PaymentReason = pr.PaymentReason,
                                      ApplicantUinType = pr.ApplicantUinTypeId,
                                      ApplicantUin = pr.ApplicantUin,
                                      ApplicantName = pr.ApplicantName,
                                      PaymentReferenceType = pr.PaymentReferenceType,
                                      PaymentReferenceNumber = pr.PaymentReferenceNumber,
                                      PaymentReferenceDate = pr.PaymentReferenceDate,
                                      ExpirationDate = pr.ExpirationDate,
                                      AdditionalInformation = pr.AdditionalInformation,
                                      CreateDate = pr.CreateDate
                                  }
                              }).ToList();

            int totalCount = allResults.Count();

            var selectResult = allResults.OrderByDescending(c => c.requestPaymentInfoParsedVO.CreateDate)
                                         .Skip((page - 1) * size)
                                         .Take(size)
                                         .ToList()
                                         .Select(e => new Tuple<int, RequestPaymentInfoParsedVO>(e.paymentRequestId, e.requestPaymentInfoParsedVO))
                                         .ToList();

            return new Tuple<int, List<Tuple<int, RequestPaymentInfoParsedVO>>>(totalCount, selectResult);

        }

        public Tuple<int, List<Tuple<PaymentRequestStatus?, List<RequestPaymentInfoParsedVO>>>> GetParsedPaymentRequestInfosByApplicantGrouped(string applicantUin, int pageNumber, int pageSize, PaymentRequestStatus[] paymentStatuses)
        {
            int page = pageNumber > 0 ? pageNumber : 1;
            int size = pageSize > 0 ? pageSize : 20;

            var allResults = (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                              join es in unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals es.EserviceClientId
                              where (pr.ApplicantUinTypeId == UinType.Egn || pr.ApplicantUinTypeId == UinType.Lnch)
                              && pr.ApplicantUin == applicantUin && paymentStatuses.Contains(pr.PaymentRequestStatusId)
                              select new
                              {
                                  requestPaymentInfoParsedVO = new RequestPaymentInfoParsedVO
                                  {
                                      Id = pr.PaymentRequestIdentifier,
                                      EserviceClientAisName = es.AisName,
                                      EserviceClientServiceName = es.ServiceName,
                                      EserviceClientDepartmentId = es.DepartmentId.ToString(),
                                      Status = pr.PaymentRequestStatusId,
                                      StatusChangeTime = pr.PaymentRequestStatusChangeTime,
                                      ServiceProviderName = pr.ServiceProviderName,
                                      ServiceProviderBank = pr.ServiceProviderBank,
                                      ServiceProviderBIC = pr.ServiceProviderBIC,
                                      ServiceProviderIBAN = pr.ServiceProviderIBAN,
                                      Currency = pr.Currency,
                                      PaymentTypeCode = pr.PaymentTypeCode,
                                      PaymentAmount = pr.PaymentAmount,
                                      PaymentReason = pr.PaymentReason,
                                      ApplicantUinType = pr.ApplicantUinTypeId,
                                      ApplicantUin = pr.ApplicantUin,
                                      ApplicantName = pr.ApplicantName,
                                      PaymentReferenceType = pr.PaymentReferenceType,
                                      PaymentReferenceNumber = pr.PaymentReferenceNumber,
                                      PaymentReferenceDate = pr.PaymentReferenceDate,
                                      ExpirationDate = pr.ExpirationDate,
                                      AdditionalInformation = pr.AdditionalInformation,
                                      CreateDate = pr.CreateDate
                                  }
                              }).ToList();

            int totalCount = allResults.Count();
            var selectResult = allResults.OrderByDescending(c => c.requestPaymentInfoParsedVO.CreateDate)
                                         .Skip((page - 1) * size)
                                         .Take(size)
                                         .GroupBy(e => e.requestPaymentInfoParsedVO.Status, e => e.requestPaymentInfoParsedVO, (key, g) => new Tuple<PaymentRequestStatus?, List<RequestPaymentInfoParsedVO>>(key, g.ToList()))
                                         .ToList();

            return new Tuple<int, List<Tuple<PaymentRequestStatus?, List<RequestPaymentInfoParsedVO>>>>(totalCount, selectResult);
        }

        public Tuple<int, List<Tuple<int, RequestPaymentInfoParsedVO>>> GetParsedPaymentRequestInfosByApplicant(string applicantUin, int pageNumber, int pageSize)
        {
            int page = pageNumber > 0 ? pageNumber : 1;
            int size = pageSize > 0 ? pageSize : 20;

            var allResults = (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                       join es in unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals es.EserviceClientId
                       where (pr.ApplicantUinTypeId == UinType.Egn || pr.ApplicantUinTypeId == UinType.Lnch)
                       && pr.ApplicantUin == applicantUin && pr.PaymentRequestStatusId != PaymentRequestStatus.Pending
                       select new
                       {
                           paymentRequestId = pr.PaymentRequestId,
                           requestPaymentInfoParsedVO = new RequestPaymentInfoParsedVO
                           {
                               Id = pr.PaymentRequestIdentifier,
                               EserviceClientAisName = es.AisName,
                               EserviceClientServiceName = es.ServiceName,
                               EserviceClientDepartmentId = es.DepartmentId.ToString(),
                               Status = pr.PaymentRequestStatusId,
                               StatusChangeTime = pr.PaymentRequestStatusChangeTime,
                               ServiceProviderName = pr.ServiceProviderName,
                               ServiceProviderBank = pr.ServiceProviderBank,
                               ServiceProviderBIC = pr.ServiceProviderBIC,
                               ServiceProviderIBAN = pr.ServiceProviderIBAN,
                               Currency = pr.Currency,
                               PaymentTypeCode = pr.PaymentTypeCode,
                               PaymentAmount = pr.PaymentAmount,
                               PaymentReason = pr.PaymentReason,
                               ApplicantUinType = pr.ApplicantUinTypeId,
                               ApplicantUin = pr.ApplicantUin,
                               ApplicantName = pr.ApplicantName,
                               PaymentReferenceType = pr.PaymentReferenceType,
                               PaymentReferenceNumber = pr.PaymentReferenceNumber,
                               PaymentReferenceDate = pr.PaymentReferenceDate,
                               ExpirationDate = pr.ExpirationDate,
                               AdditionalInformation = pr.AdditionalInformation,
                               CreateDate = pr.CreateDate
                           }
                       }).ToList();

            int totalCount = allResults.Count();
            var selectResult = allResults.OrderByDescending(c => c.requestPaymentInfoParsedVO.CreateDate)
                                         .Skip((page - 1) * size)
                                         .Take(size)
                                         .ToList()
                                         .Select(e => new Tuple<int, RequestPaymentInfoParsedVO>(e.paymentRequestId, e.requestPaymentInfoParsedVO))
                                         .ToList();

            return new Tuple<int, List<Tuple<int, RequestPaymentInfoParsedVO>>>(totalCount, selectResult);
        }

        public List<RequestEikInfoVO> GetServiceClientByEik(string eiknumber) 
        {
            var result = from dep in this.unitOfWork.DbContext.Set<Department>()
                         join esc in this.unitOfWork.DbContext.Set<EserviceClient>() on dep.DepartmentId equals esc.DepartmentId
                         where dep.UniqueIdentificationNumber == eiknumber && dep.IsActive
                         select
                             new RequestEikInfoVO()
                             {
                                 DepartmentId = dep.DepartmentId,
                                 Name = esc.ServiceName,
                                 UniqueIdentificationNumber = dep.UniqueIdentificationNumber,
                                 IsActive = dep.IsActive,
                                 EserviceClientId = esc.EserviceClientId,
                                 AisName = esc.AisName,
                                 AccountBank = esc.AccountBank,
                                 AccountBIC = esc.AccountBIC,
                                 AccountIBAN = esc.AccountIBAN
                             };
             return result.ToList();
            }

        public List<Tuple<int, RequestEikInfoVO>> GetParsedRequestInfoByEik(string eiknumber)
        {
            return (from dep in this.unitOfWork.DbContext.Set<Department>()
                    join esc in this.unitOfWork.DbContext.Set<EserviceClient>() on dep.DepartmentId equals esc.DepartmentId
                    where dep.UniqueIdentificationNumber == eiknumber && dep.IsActive 
                    select new
                    {
                        DepartmentId = dep.DepartmentId,
                        requestEikInfoVO = new RequestEikInfoVO()
                        {
                            DepartmentId = dep.DepartmentId,
                            Name = dep.Name,
                            UniqueIdentificationNumber = dep.UniqueIdentificationNumber,
                            IsActive = dep.IsActive,
                            EserviceClientId = esc.EserviceClientId,
                            AisName = esc.AisName,
                            AccountBank = esc.AccountBank,
                            AccountBIC = esc.AccountBIC,
                            AccountIBAN = esc.AccountIBAN
                        }

                    }).ToList()
                    .Select( e => new Tuple<int, RequestEikInfoVO>(e.DepartmentId, e.requestEikInfoVO))
                    .ToList();
        }

        public List<Tuple<int, RequestRefidInfoVO>> GetParsedRequestInfoByRefid(int refid, string clientId)
        {

            var result = (from dr in unitOfWork.DbContext.Set<DistributionRevenue>()
                          join drp in unitOfWork.DbContext.Set<DistributionRevenuePayment>() on dr.DistributionRevenueId equals drp.DistributionRevenueId
                          join bt in unitOfWork.DbContext.Set<BoricaTransaction>() on drp.BoricaTransactionId equals bt.BoricaTransactionId
                          join es in unitOfWork.DbContext.Set<EserviceClient>() on drp.EserviceClientId equals es.EserviceClientId
                          join depart in unitOfWork.DbContext.Set<Department>() on es.DepartmentId equals depart.DepartmentId
                          where dr.DistributionRevenueId == refid && es.ClientId == clientId
                          select new
                          {
                              BoricaTransactionId = bt.BoricaTransactionId,
                              DistributionRevenueId = dr.DistributionRevenueId,
                              requestRefidInfoVO = new RequestRefidInfoVO()
                              {
                                  DistributionRevenueId = dr.DistributionRevenueId,
                                  CreatedAt = dr.CreatedAt,
                                  Order = bt.Order,
                                  Amount = bt.Amount,
                                  TransactionDate = bt.TransactionDate,
                                  Rrn = bt.Rrn,
                                  PaymentInfo =
                                  bt.PaymentRequests.Select(pr => new RequestPaymentInfoParsedVO
                                  {
                                      Id = pr.PaymentRequestId.ToString(),
                                      Status = pr.PaymentRequestStatusId,
                                      StatusChangeTime = pr.PaymentRequestStatusChangeTime,
                                      ServiceProviderName = pr.ServiceProviderName,
                                      ServiceProviderBank = pr.ServiceProviderBank,
                                      ServiceProviderBIC = pr.ServiceProviderBIC,
                                      ServiceProviderIBAN = pr.ServiceProviderIBAN,
                                      Currency = pr.Currency,
                                      PaymentTypeCode = pr.PaymentTypeCode,
                                      PaymentAmount = pr.PaymentAmount,
                                      PaymentReason = pr.PaymentReason,
                                      ApplicantUinType = pr.ApplicantUinTypeId,
                                      ApplicantUin = pr.ApplicantUin,
                                      ApplicantName = pr.ApplicantName,
                                      PaymentReferenceType = pr.PaymentReferenceType,
                                      PaymentReferenceNumber = pr.PaymentReferenceNumber,
                                      PaymentReferenceDate = pr.PaymentReferenceDate,
                                      ExpirationDate = pr.ExpirationDate,
                                      AdditionalInformation = pr.AdditionalInformation,
                                      CreateDate = pr.CreateDate,
                                      EserviceClientAisName = es.AisName,
                                      EserviceClientServiceName = es.ServiceName,
                                      ClientId = es.ClientId,
                                      EserviceClientDepartmentId = es.DepartmentId.ToString()
                                  }),
                                  DepartmentName = depart.Name,
                                  DepartmentUniqueIdentificationNumber = depart.UniqueIdentificationNumber
                              }
                          })
                    .GroupBy(g => g.BoricaTransactionId)
                    .Select(s => s.FirstOrDefault()).ToList()
                    .Select(e => new Tuple<int, RequestRefidInfoVO>(e.DistributionRevenueId, e.requestRefidInfoVO))
                    .ToList();

            decimal sum = decimal.Zero;
            foreach(var item in result)
            {
                sum += item.Item2.Amount;
            }
            foreach (var item in result)
            {
                item.Item2.TotalSum = sum;
            }

            return result;
        }

        public int GetNumberOfPaymentsByMonth(string request_type, string month)
        {
            int result = 0;
            try
            {
                if (string.IsNullOrEmpty(month) == false || string.IsNullOrWhiteSpace(month))
                {
                    DateTime date = DateTime.Parse(month);
                    int daysInMonth = DateTime.DaysInMonth(date.Year, date.Month);
                    DateTime startDate = new DateTime(date.Year, date.Month, 1);
                    DateTime endDate = new DateTime(date.Year, date.Month, daysInMonth);

                    if (request_type == "request")
                    {
                        result = this.unitOfWork.DbContext.Set<PaymentRequest>().
                            Count(pr => (pr.CreateDate >= startDate && pr.CreateDate <= endDate));
                    }
                    else if (request_type == "payments")
                    {
                        result = this.unitOfWork.DbContext.Set<PaymentRequest>().
                            Count(pr => (pr.CreateDate >= startDate && pr.CreateDate <= endDate)
                            && pr.PaymentRequestStatusId == PaymentRequestStatus.Paid);
                    }
                }
                else
                {
                    DateTime startDate = new DateTime(DateTime.Now.Year, 1, 1);
                    DateTime endDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                    for (var m = startDate.Month; m <= endDate.Month; m++)
                    {

                        if (request_type == "request")
                        {
                            result = this.unitOfWork.DbContext.Set<PaymentRequest>().
                                Count(pr => (pr.CreateDate >= startDate && pr.CreateDate <= endDate));
                        }
                        else if (request_type == "payments")
                        {
                            result = this.unitOfWork.DbContext.Set<PaymentRequest>().
                                Count(pr => (pr.CreateDate >= startDate && pr.CreateDate <= endDate)
                                && pr.PaymentRequestStatusId == PaymentRequestStatus.Paid);
                        }
                    }
                }
            }
            
            catch (Exception ex)
            {

            }

            return result;
        }

        public int GetNumberOfPaymentsFromDateToDate(string request_type, string startDate, string endDate)
        {
            int result = 0;
            try
            {
                DateTime startdate = DateTime.Parse(startDate);
                DateTime enddate = DateTime.Parse(endDate);

                if (request_type == "request")
                {
                    result = this.unitOfWork.DbContext.Set<PaymentRequest>().
                        Count(pr => (pr.CreateDate >= startdate && pr.CreateDate <= enddate));
                }
                else if (request_type == "payments")
                {
                    result = this.unitOfWork.DbContext.Set<PaymentRequest>().
                        Count(pr => (pr.CreateDate >= startdate && pr.CreateDate <= enddate)
                        && pr.PaymentRequestStatusId == PaymentRequestStatus.Paid);
                }
            }

            catch (Exception ex)
            { }

            return result;
        }

        public object GetNumberOfVposePayCvposPaidRequestsFromDateToDate(string start_date, string end_date)
        {
            object result = new object();

            try
            {
                DateTime startdate = DateTime.Parse(start_date);
                DateTime enddate = DateTime.Parse(end_date);

                var vposres = this.unitOfWork.DbContext.Set<VposFiBankRequest>().
                        Count(fb => (fb.CreateDate >= startdate && fb.CreateDate <= enddate && fb.IsPaymentSuccessful == true))
                        + this.unitOfWork.DbContext.Set<VposDskEcommRequest>().
                        Count(ds => (ds.CreateDate >= startdate && ds.CreateDate <= enddate && ds.IsPaymentSuccessful == true));

                var epayres = this.unitOfWork.DbContext.Set<VposEpayRequest>().
                        Count(ep => (ep.CreateDate >= startdate && ep.CreateDate <= enddate && ep.IsPaymentSuccessful == true));

                var cvposres = this.unitOfWork.DbContext.Set<BoricaTransaction>().
                        Count(br => (br.TransactionDate >= startdate && br.TransactionDate <= enddate && (br.TransactionStatusId == 2 ||
                        br.TransactionStatusId == 5)));

                var res = new
                {
                    vpos = new
                    {
                        from = $"{start_date} to {end_date}",
                        value = vposres
                    }, 
                    epay = new
                    {
                        from = $"{start_date} to {end_date}",
                        value = epayres
                    },
                    cvpos = new
                    {
                        from = $"{start_date} to {end_date}",
                        value = cvposres
                    }     
                };

                result = res;
            }

            catch (Exception ex)
            { }

            return result;
        }

        public object GetAdministrationsWithTheHighestNumberOfPaymentsFromDateToDate(string start_date, string end_date)
        {
            object result = new object();

            try
            {
                DateTime startdate = DateTime.Parse(start_date);
                DateTime enddate = DateTime.Parse(end_date);
                var resadmin = (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                           join esc in unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals esc.EserviceClientId
                           join dep in unitOfWork.DbContext.Set<Department>() on esc.DepartmentId equals dep.DepartmentId

                           where pr.CreateDate >= startdate && pr.CreateDate <= enddate && pr.PaymentRequestStatusId == PaymentRequestStatus.Paid
                           select new
                           {
                               AdministrationId = pr.EserviceClientId,
                               ApplicantName = pr.ApplicantName,
                               AdministrationName = dep.Name
                           }).ToList().GroupBy(r => r.AdministrationId).OrderByDescending(r => r.Count()).ToList().FirstOrDefault();

                if(resadmin.Count() > 0)
                {
                    result = new
                    {
                        from = $"{start_date} to {end_date}",
                        Administration = resadmin.FirstOrDefault().AdministrationName,
                        value = resadmin.Count()
                    };
                }
            }

            
            catch (Exception ex)
            { }

            return result;
        }

        public Tuple<int, RequestInfoVO> GetRequestInfoByAccessCode(string accessCode)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    join prx in this.unitOfWork.DbContext.Set<PaymentRequestXml>() on pr.PaymentRequestXmlId equals prx.PaymentRequestXmlId
                    where pr.PaymentRequestAccessCode == accessCode
                    select new
                    {
                        paymentRequestId = pr.PaymentRequestId,
                        requestInfoVO = new RequestInfoVO
                        {
                            Id = pr.PaymentRequestIdentifier,
                            Status = pr.PaymentRequestStatusId,
                            ChangeTime = pr.PaymentRequestStatusChangeTime,
                            RequestXml = prx.RequestContent
                        }
                    })
                    .ToList()
                    .Select(e => new Tuple<int, RequestInfoVO>(e.paymentRequestId, e.requestInfoVO))
                    .SingleOrDefault();
        }

        public Tuple<int, RequestInfoParsedVO> GetParsedRequestInfoByAccessCode(string accessCode)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    where pr.PaymentRequestAccessCode == accessCode
                    select new
                    {
                        paymentRequestId = pr.PaymentRequestId,
                        requestInfoParsedVO = new RequestInfoParsedVO
                        {
                            Id = pr.PaymentRequestIdentifier,
                            Status = pr.PaymentRequestStatusId,
                            StatusChangeTime = pr.PaymentRequestStatusChangeTime,
                            ServiceProviderName = pr.ServiceProviderName,
                            ServiceProviderBank = pr.ServiceProviderBank,
                            ServiceProviderBIC = pr.ServiceProviderBIC,
                            ServiceProviderIBAN = pr.ServiceProviderIBAN,
                            Currency = pr.Currency,
                            PaymentTypeCode = pr.PaymentTypeCode,
                            PaymentAmount = pr.PaymentAmount,
                            PaymentReason = pr.PaymentReason,
                            ApplicantUinType = pr.ApplicantUinTypeId,
                            ApplicantUin = pr.ApplicantUin,
                            ApplicantName = pr.ApplicantName,
                            PaymentReferenceType = pr.PaymentReferenceType,
                            PaymentReferenceNumber = pr.PaymentReferenceNumber,
                            PaymentReferenceDate = pr.PaymentReferenceDate,
                            ExpirationDate = pr.ExpirationDate,
                            AdditionalInformation = pr.AdditionalInformation,
                            CreateDate = pr.CreateDate
                        }
                    })
                   .ToList()
                   .Select(e => new Tuple<int, RequestInfoParsedVO>(e.paymentRequestId, e.requestInfoParsedVO))
                   .SingleOrDefault();
        }

        public bool IsValidRequestWithKeyDataExist(string serviceProviderIBAN, string paymentReferenceNumber, DateTime paymentReferenceDate)
        {
            var isRequestExist = false;

            // we order today's requests by CreateDate so that we get
            // the LAST added payment request on top of the list
            var existingRequest = this.unitOfWork.DbContext.Set<PaymentRequest>().Include(bt=> bt.BoricaTransactions)
                .Where(e =>
                    e.ServiceProviderIBAN == serviceProviderIBAN &&
                    e.PaymentReferenceNumber == paymentReferenceNumber &&
                    (e.PaymentReferenceDate == paymentReferenceDate))
                .OrderByDescending(e => e.CreateDate)
                .FirstOrDefault();

            if (existingRequest != null)
            {
                if (existingRequest.PayOrder != null)
                {
                    if (existingRequest.ExpirationDate <= DateTime.Now)
                    {
                        // exists && MDT && expired => INSERT
                        isRequestExist = false;

                        if (existingRequest.PaymentRequestStatusId == PaymentRequestStatus.Authorized
                            || existingRequest.PaymentRequestStatusId == PaymentRequestStatus.Ordered
                            || existingRequest.PaymentRequestStatusId == PaymentRequestStatus.Paid
                            || existingRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess)
                        {
                            // exists && MDT && expired && status in (AUTHORIZED, ORDERED, PAID, INPROCESS) => UPDATE
                            isRequestExist = true;
                        }
                        var transactionExists =
                            (from bt in this.unitOfWork.DbContext.Set<BoricaTransaction>()
                             from pr in bt.PaymentRequests
                             where pr.PaymentRequestId == existingRequest.PaymentRequestId
                                && bt.TransactionStatusId != (int)BoricaTransactionStatusEnum.Canceled
                             select bt)
                            .Any();

                        if (transactionExists)
                        {
                            // exists && MDT && expired && Borica transaction exists => UPDATE
                            isRequestExist = true;
                        }
                    }
                    else
                    {
                        // exists && MDT && NOT expired  => UPDATE
                        isRequestExist = true;
                    }
                }
                else
                {
                    // exists && NOT MDT => UPDATE
                    isRequestExist = true;
                }
            } 
            else 
            {
                // NOT exists => INSERT
            }


            return isRequestExist;
        }

        public bool IsClientAuthorizedToAccessRequests(string clientId, List<string> paymentRequestIdentifiers, bool authorizeIfClientIdIsPaymentInitiator = false)
        {
            if (authorizeIfClientIdIsPaymentInitiator)
            {
                return !(from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                         join ini in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.InitiatorId equals ini.EserviceClientId
                         join esc in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals esc.EserviceClientId
                         where paymentRequestIdentifiers.Contains(pr.PaymentRequestIdentifier) && esc.ClientId != clientId && ini.ClientId != clientId
                         select pr)
                   .Any();
            }
            return !(from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                        join esc in this.unitOfWork.DbContext.Set<EserviceClient>() on pr.EserviceClientId equals esc.EserviceClientId
                        where paymentRequestIdentifiers.Contains(pr.PaymentRequestIdentifier) && esc.ClientId != clientId
                        select pr)
                .Any();
        }

        public PaymentRequest GetPaymentRequestByAisPaymentId(int eserviceClientId, string aisPaymentId)
        {
            return this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(e => e.EserviceClientId == eserviceClientId && e.AisPaymentId == aisPaymentId)
                .SingleOrDefault();
        }

        public List<ObligationType> GetAllObligationTypes()
        {
            return this.unitOfWork.DbContext.Set<ObligationType>().ToList();
        }
    }
}
