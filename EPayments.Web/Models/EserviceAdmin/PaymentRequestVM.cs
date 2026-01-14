using EPayments.Web.DataObjects;
using System;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class PaymentRequestVM
    {
        public string Title { get; set; } = "Създай нова заявка за плащане";

        public PaymentRequestDO PaymentRequest { get; set; } = new PaymentRequestDO();

        public string BackButtonText = "Назад";
    }
}