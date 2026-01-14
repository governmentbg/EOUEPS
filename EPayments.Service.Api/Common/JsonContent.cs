using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Common
{
    public class JsonContent : HttpContent
    {
        private readonly JToken JTokenValue;

        public JsonContent(JToken value)
        {
            this.JTokenValue = value;
            Headers.ContentType = new MediaTypeHeaderValue("application/json")
            {
                CharSet = "utf-8"
            };
        }

        protected override Task SerializeToStreamAsync(Stream stream,
            TransportContext context)
        {
            var jsonWriter = new JsonTextWriter(new StreamWriter(stream))
            {
                Formatting = Formatting.Indented
            };

            this.JTokenValue.WriteTo(jsonWriter);
            jsonWriter.Flush();

            return Task.FromResult<object>(null);
        }

        protected override bool TryComputeLength(out long length)
        {
            length = -1;

            return false;
        }
    }
}
