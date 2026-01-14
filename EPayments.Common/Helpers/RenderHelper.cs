using RazorEngine;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RazorEngine.Templating;
using EPayments.Common.Helpers;

namespace EPayments.Common.Helpers
{
    public class RenderHelper
    {
        private static object syncRoot = new object();

        public static byte[] RenderPdf(string viewTemplate, string modelHtml)
        {
            string html = RenderHelper.RenderHtmlByMvcView(viewTemplate, (object)modelHtml);
            string mimeType = MimeTypeFileExtension.MIME_TEXT_HTML;

            return PdfConvertManager.Convert(html, ref mimeType);
        }

        public static string RenderHtmlByMvcView(string viewTemplate, object model)
        {
            string templateHtml = File.ReadAllText(System.Web.Hosting.HostingEnvironment.MapPath(viewTemplate));

            return RenderView(templateHtml, viewTemplate, model);
        }

        public static string RenderHtmlByRazorTemplate(string templateName, string templatePath, object model)
        {
            string templateHtml = File.ReadAllText(templatePath);

            return RenderView(templateHtml, templateName, model);
        }

        private static string RenderView(string templateHtml, string templateName, object model)
        {
            if (!Engine.Razor.IsTemplateCached(templateName, model.GetType()))
            {
                    if (!Engine.Razor.IsTemplateCached(templateName, model.GetType()))
                    {
                        Engine.Razor.Compile(templateHtml, templateName, model.GetType());
                        Engine.Razor.AddTemplate(templateName, templateHtml);
                    }
            }

            return Engine.Razor.Run(templateName, model.GetType(), model, null);
        }
    }
}
