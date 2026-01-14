using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Implementations;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Common
{
    public class DocumentHelper
    {
        public static Tuple<UinType, string, string> GetUinTypeUinAndName(R_0009_000015.ElectronicServiceRecipient electronicServiceRecipient)
        {
            UinType? applicantUinType = null;
            string applicantUin = null;
            string applicantName = null;

            if (electronicServiceRecipient.Person != null)
            {
                if (!String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Identifier.EGN))
                {
                    applicantUinType = UinType.Egn;
                    applicantUin = electronicServiceRecipient.Person.Identifier.EGN;
                }
                else if (!String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Identifier.LNCh))
                {
                    applicantUinType = UinType.Lnch;
                    applicantUin = electronicServiceRecipient.Person.Identifier.LNCh;
                }

                applicantName = Formatter.FormatName(
                    electronicServiceRecipient.Person.Names.First,
                    electronicServiceRecipient.Person.Names.Middle,
                    electronicServiceRecipient.Person.Names.Last);
            }
            else if (electronicServiceRecipient.Entity != null)
            {
                applicantUinType = UinType.Bulstat;
                applicantUin = electronicServiceRecipient.Entity.Identifier;
                applicantName = electronicServiceRecipient.Entity.Name;
            }

            if (!applicantUinType.HasValue || String.IsNullOrWhiteSpace(applicantUin) || String.IsNullOrWhiteSpace(applicantName))
            {
                /**/
                throw new Exception("Missing data for electronicServiceApplicant");
            }

            return new Tuple<UinType, string, string>(applicantUinType.Value, applicantUin, applicantName);
        }

        public static List<string> ValidatePaymentRequestData(IApiRepository apiRepository, R_10046.PaymentRequest paymentRequest, int eserviceClientId, string aisPaymentId, IPaymentRequestRepository paymentRequestRepository, out PaymentRequest existingRequest)
        {
            existingRequest = null;

            List<string> errors = new List<string>();

            bool hasValidBankAccountIBAN = false;
            bool hasValidPaymentReferenceNumber = false;
            bool hasValidPaymentReferenceDate = false;
            bool hasValidExpirationDate = false;
            ObligationType obligationType = null;

            if (!string.IsNullOrEmpty(paymentRequest.ObligationType))
            {
                var isParsed = int.TryParse(paymentRequest.ObligationType, out int obligationTypeId);
                obligationType = apiRepository.GetAllObligationTypes().FirstOrDefault(o => o.ObligationTypeId == obligationTypeId);
                if (!isParsed || obligationType == null)
                {
                    errors.Add("Element Obligation type must be valid integer.");
                }
            }

            if (paymentRequest.ElectronicServiceProviderBasicData == null)
            {
                errors.Add("Element 'ElectronicServiceProviderBasicData' is required.");
            }
            else
            {
                if (paymentRequest.ElectronicServiceProviderBasicData.EntityBasicData == null)
                {
                    errors.Add("Element 'ElectronicServiceProviderBasicData.EntityBasicData' is required.");
                }
                else
                {
                    if (String.IsNullOrWhiteSpace(paymentRequest.ElectronicServiceProviderBasicData.EntityBasicData.Name))
                    {
                        errors.Add("Element 'ElectronicServiceProviderBasicData.EntityBasicData.Name' is required.");
                    }
                }
            }

            if (paymentRequest.ElectronicServiceProviderBankAccount == null)
            {
                errors.Add("Element 'ElectronicServiceProviderBankAccount' is required.");
            }
            else
            {
                if (paymentRequest.ElectronicServiceProviderBankAccount.EntityBasicData == null)
                {
                    errors.Add("Element 'ElectronicServiceProviderBankAccount.EntityBasicData' is required.");
                }
                else
                {
                    if (String.IsNullOrWhiteSpace(paymentRequest.ElectronicServiceProviderBankAccount.EntityBasicData.Name))
                    {
                        errors.Add("Element 'ElectronicServiceProviderBankAccount.EntityBasicData.Name' is required.");
                    }
                }

                if (String.IsNullOrWhiteSpace(paymentRequest.ElectronicServiceProviderBankAccount.BIC))
                {
                    errors.Add("Element 'ElectronicServiceProviderBankAccount.BIC' is required.");
                }

                if (String.IsNullOrWhiteSpace(paymentRequest.ElectronicServiceProviderBankAccount.IBAN))
                {
                    errors.Add("Element 'ElectronicServiceProviderBankAccount.IBAN' is required.");
                }
                else
                {
                    hasValidBankAccountIBAN = true;
                }
            }

            if (String.IsNullOrWhiteSpace(paymentRequest.Currency))
            {
                errors.Add("Element 'Currency' is required.");
            }

            if (String.IsNullOrWhiteSpace(paymentRequest.PaymentReason))
            {
                errors.Add("Element 'PaymentReason' is required.");
            }

            if (paymentRequest.PaymentReference == null)
            {
                errors.Add("Element 'PaymentReference' is required.");
            }
            else
            {
                if (String.IsNullOrWhiteSpace(paymentRequest.PaymentReference.PaymentReferenceNumber))
                {
                    errors.Add("Element 'PaymentReference.PaymentReferenceNumber' is required.");
                }

                int obligationTypeId = paymentRequest.ObligationType != null ? int.Parse(paymentRequest.ObligationType) : 1;
                var allObligationTypes = apiRepository.GetAllObligationTypes();
                obligationType = allObligationTypes.Find(x => x.ObligationTypeId == obligationTypeId);
                if (obligationType == null)
                {
                    errors.Add("Element 'PaymentReference.ObligationType' is not present in the Database.");
                }
                else
                {
                    if (paymentRequest.PaymentReference.PaymentReferenceNumber != null)
                    {
                        if (obligationType.AlgorithmId == 0)
                        {
                            if (paymentRequestRepository.IsPaymentReferenceNumberExists(paymentRequest.PaymentReference.PaymentReferenceNumber))
                            {
                                errors.Add("Element 'PaymentReference.PaymentReferenceNumber' is already present in the Database.");
                            }
                            else
                            {
                                hasValidPaymentReferenceNumber = true;
                            }
                        }
                        else if (obligationType.AlgorithmId == 1)
                        {
                            if (paymentRequestRepository.IsPaymentReferenceNumberExistsForAlgorithm(0, paymentRequest.PaymentReference.PaymentReferenceNumber))
                            {
                                errors.Add("Element 'PaymentReference.PaymentReferenceNumber' is already present in the Database for Algorithm 0.");
                            }
                            else
                            {
                                hasValidPaymentReferenceNumber = true;
                            }
                        }
                        else if (obligationType.AlgorithmId == 2)
                        {
                            string paymentReferenceNumber = paymentRequest.PaymentReference.PaymentReferenceNumber.Trim();

                            List<PaymentRequest> payments = paymentRequestRepository.GetPendingPaymentRequestByPaymentReferenceNumberAndAisPaymentId(paymentReferenceNumber, aisPaymentId);
                            string errorMessage = "Element 'PaymentReference.PaymentReferenceNumber' validation failed. ";
                            bool statusPaidOrInProcess = false;
                            if (payments != null)
                            {
                                if (payments.Find(x => x.PaymentRequestStatusId == PaymentRequestStatus.Paid) != null)
                                {
                                    errorMessage += "A payment with this referenceNumber and status paid exists.";
                                    statusPaidOrInProcess = true;
                                }
                                else if (payments.Find(x => x.PaymentRequestStatusId == PaymentRequestStatus.InProcess) != null)
                                {
                                    errorMessage += "A payment with this referenceNumber and status inProcess exists.";
                                    statusPaidOrInProcess = true;
                                }
                                else
                                { 
                                    foreach (PaymentRequest payment in payments.Where(x => x.PaymentRequestStatusId != PaymentRequestStatus.Expired))
                                    {
                                        payment.PaymentRequestStatusId = PaymentRequestStatus.Expired;
                                    }
                                }
                            }

                            if (statusPaidOrInProcess)
                            {
                                errors.Add(errorMessage);
                            }
                            else
                            { 
                                hasValidPaymentReferenceNumber = true;
                            }
                        }
                        else
                        {
                            errors.Add("Element 'PaymentReference.PaymentReferenceNumber' validation not implemented for this Algorithm.");
                        }
                    }
                    else
                    {
                        hasValidPaymentReferenceNumber = true;
                    }
                }


                if (!paymentRequest.PaymentReference.PaymentReferenceDate.HasValue)
                {
                    errors.Add("Element 'PaymentReference.PaymentReferenceDate' is required.");
                }
                else
                {
                    hasValidPaymentReferenceDate = true;
                }
            }

            if (!paymentRequest.PaymentRequestExpirationDate.HasValue)
            {
                errors.Add("Element 'PaymentRequestExpirationDate' is required.");
            }
            else
            {
                if (paymentRequest.PaymentRequestExpirationDate.Value <= DateTime.Now)
                {
                    errors.Add("Value of 'PaymentRequestExpirationDate' should be greater than current date.");
                }
                else
                {
                    hasValidExpirationDate = true;
                }
            }

            if (hasValidBankAccountIBAN && hasValidPaymentReferenceNumber && hasValidPaymentReferenceDate && hasValidExpirationDate)
            {
                if (String.IsNullOrWhiteSpace(aisPaymentId))
                {
                    if (!(obligationType != null && obligationType.AlgorithmId == 2))
                    {
                        bool isRequestExist = apiRepository.IsValidRequestWithKeyDataExist(
                        paymentRequest.ElectronicServiceProviderBankAccount.IBAN.Trim(),
                        paymentRequest.PaymentReference.PaymentReferenceNumber.Trim(),
                        paymentRequest.PaymentReference.PaymentReferenceDate.Value.Date);

                        if (isRequestExist)
                        {
                            errors.Add("Заявка за плащене със същия референтен номер вече съществува!");
                        }
                    }
                }
                else
                {
                    if (!(obligationType != null && obligationType.AlgorithmId == 2))
                    {
                        existingRequest = apiRepository.GetPaymentRequestByAisPaymentId(eserviceClientId, aisPaymentId);
                        if (existingRequest != null)
                        {
                            if (existingRequest.PaymentRequestStatusId != PaymentRequestStatus.Pending && !existingRequest.PayOrder.HasValue)
                            {
                                errors.Add(String.Format("Payment request with corresponding aisPaymentId has status '{0}'. Payment request replacing is allowed only when request status is 'Pending'.", existingRequest.PaymentRequestStatusId.ToString()));
                            }

                            /*
                            if (existingRequest.PaymentRequestStatusId != PaymentRequestStatus.InProcess && 
                                existingRequest.PaymentRequestStatusId != PaymentRequestStatus.Pending &&
                                existingRequest.PayOrder.HasValue)
                             */

                            if (existingRequest.PaymentRequestStatusId == PaymentRequestStatus.InProcess && existingRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending && existingRequest.PayOrder.HasValue)
                            {
                                errors.Add(String.Format("Payment request with corresponding aisPaymentId has status '{0}'. Payment request replacing is allowed only when request status is 'Pending'.", existingRequest.PaymentRequestStatusId.ToString()));
                            }
                        }
                    }
                }
            }

            if (paymentRequest.ElectronicServiceRecipient == null)
            {
                errors.Add("Element 'ElectronicServiceRecipient' is required.");
            }
            else
            {
                var electronicServiceRecipient = paymentRequest.ElectronicServiceRecipient;

                bool hasUinTypeSpecified = false;
                bool hasUinSpecified = false;
                bool hasNameSpecified = false;

                if (electronicServiceRecipient.Person != null)
                {
                    hasUinTypeSpecified = true;

                    if (electronicServiceRecipient.Person.Identifier != null)
                    {
                        if (!String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Identifier.EGN) ||
                            !String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Identifier.LNCh))
                        {
                            hasUinSpecified = true;
                        }
                    }

                    if (electronicServiceRecipient.Person.Names != null)
                    {
                        if (!String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Names.First) ||
                            !String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Names.Middle) ||
                            !String.IsNullOrWhiteSpace(electronicServiceRecipient.Person.Names.Last))
                        {
                            hasNameSpecified = true;
                        }
                    }
                }
                else if (electronicServiceRecipient.Entity != null)
                {
                    hasUinTypeSpecified = true;

                    if (!String.IsNullOrWhiteSpace(electronicServiceRecipient.Entity.Identifier))
                    {
                        hasUinSpecified = true;
                    }

                    if (!String.IsNullOrWhiteSpace(electronicServiceRecipient.Entity.Name))
                    {
                        hasNameSpecified = true;
                    }
                }

                if (!hasUinTypeSpecified)
                {
                    errors.Add("Element 'ElectronicServiceRecipient' (Person or Entity) is required.");
                }

                if (!hasUinSpecified)
                {
                    errors.Add("Recipient identifier in 'ElectronicServiceRecipient' (Person or Entity) is required.");
                }

                if (!hasNameSpecified)
                {
                    errors.Add("Recipient name in 'ElectronicServiceRecipient' (Person or Entity) is required.");
                }
            }

            return errors;
        }
    }
}
