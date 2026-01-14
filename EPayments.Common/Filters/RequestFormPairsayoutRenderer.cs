using NLog;
using NLog.LayoutRenderers;
using System;
using System.Text;
using System.Web;

namespace EPayments.Common.Filters
{
    [LayoutRenderer("requestFormPairs")]
    public class RequestFormPairsayoutRenderer : LayoutRenderer
    {
        private static readonly string RequestFormPairsKey = "__NLogRequestFormPairsKey__";

        protected override void Append(StringBuilder builder, LogEventInfo ev)
        {
            HttpContext context = HttpContext.Current;
            
            if (context == null)
            {
                builder.Append(Guid.Empty);
            }
            else
            {
                if (!context.Items.Contains(RequestFormPairsKey))
                {
                    var form = new StringBuilder();
            
                    if (context.Handler != null && context.Request.Form != null && context.Request.Form.Keys.Count > 0)
                    {
                        foreach (var key in context.Request.Form.AllKeys)
                        {
                            form.AppendFormat("{0}:{1}{2}", key, context.Request.Form[key], Environment.NewLine);
                        }
                    }
            
                    context.Items.Add(RequestFormPairsKey, form.ToString());
                }
            
                builder.Append(context.Items[RequestFormPairsKey]);
            }
        }
    }
}
