using log4net;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http.Filters;

namespace EPayments.Service.Api.Common.CustomExceptions
{
    public class CustomExceptionFilter : ExceptionFilterAttribute
    {
        public override void OnException(HttpActionExecutedContext context)
        {
            var logger = LogManager.GetLogger(typeof(CustomExceptionFilter));
            logger.Error(context.Exception);

            if (context.Exception is CustomServiceException)
            {
                JObject json = (JObject)JToken.FromObject(new { Message = context.Exception.Message });

                context.Response = new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new JsonContent(json)
                };
            }
        }
    }
}
