using System;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Web.Security;
using System.Text;
using System.Web;
using System.Collections.Specialized;
using System.Configuration;
using EPayments.CertificateUtils;
using EPayments.CertificateUtils.Extractor;
using System.Web.UI;
using System.Web.Mvc;
using EPayments.Web.Common;
using EPayments.Common;

namespace EPayments.Web.SSL
{
    public class Cert : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            byte[] certificate = null;

            if (AppSettings.EPaymentsWeb_UseFakeCertificate)
            {
                certificate = Convert.FromBase64String(AppSettings.EPaymentsWeb_FakeCertificateBase64Data);
            }
            else
            {
                certificate = Request.ClientCertificate.Certificate;
            }

            Session[SessionKeys.CertificateRawContent] = certificate;

#if DEBUG
            Response.Redirect("~/Account/Login", false);
            Context.ApplicationInstance.CompleteRequest();
#else
        Response.Redirect("~/Account/Login");
#endif

        }

        private void WriteException(StringBuilder stringBuilder, Exception exception)
        {
            stringBuilder.AppendFormat("Exception type: {0}\n", exception.GetType().FullName);
            stringBuilder.AppendFormat("Message: {0}\n", exception.Message);
            stringBuilder.AppendFormat("Stack trace:\n{0}\n", exception.StackTrace);
            if (exception.InnerException != null)
            {
                stringBuilder.AppendFormat("\n\nInner Exception:\n");
                WriteException(stringBuilder, exception.InnerException);
            }
        }

        private void WriteRequest(StringBuilder stringBuilder, HttpRequest request)
        {
            stringBuilder.AppendLine("Request data:");
            stringBuilder.AppendFormat("UserHostAddress: {0}\n", request.UserHostAddress);
            stringBuilder.AppendFormat("RawUrl: {0}\n", request.RawUrl);
            stringBuilder.AppendFormat("Form: {0}\n", GetFormString(request.Form));
            stringBuilder.AppendFormat("UserAgent: {0}\n", request.UserAgent);
        }

        private string GetFormString(NameValueCollection form)
        {
            int i = 0;
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string key in form.AllKeys)
            {
                string value = form[key];

                if (key.ToLower().Contains("password"))
                    continue;

                if (i == 0)
                {
                    stringBuilder.AppendFormat("{0}={1}", key, value);
                }
                else
                {
                    stringBuilder.AppendFormat("&{0}={1}", key, value);
                }

                i++;
            }

            return stringBuilder.ToString();
        }
    }
}