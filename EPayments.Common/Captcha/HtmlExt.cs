using System;
using System.Text;
using System.Web;
using System.Web.Caching;
using System.Web.Mvc;

namespace EPayments.Common.Captcha
{
    public static class HtmlExt
    {
        /// <summary>
        /// Generates the captcha image.
        /// </summary>
        /// <param name="helper">The helper.</param>
        /// <param name="height">The height.</param>
        /// <param name="width">The width.</param>
        /// <returns>
        /// Returns the <see cref="Uri"/> for the generated <see cref="CaptchaImage"/>.
        /// </returns>
        public static MvcHtmlString CaptchaImage(this HtmlHelper helper, CaptchaDifficultyLevel difficultyLevel, int height, int width, string classValue = null, string styleValue = null)
        {
            CaptchaImage image = new CaptchaImage(difficultyLevel, height, width);

            HttpRuntime.Cache.Add(
                    image.UniqueId,
                    image,
                    null,
                    DateTime.Now.AddSeconds(Captcha.CaptchaImage.CacheTimeOut),
                    Cache.NoSlidingExpiration,
                    CacheItemPriority.NotRemovable,
                    null);

            StringBuilder stringBuilder = new StringBuilder(256);
            stringBuilder.Append("<input type=\"hidden\" name=\"validation-guid\" value=\"");
            stringBuilder.Append(image.UniqueId);
            stringBuilder.Append("\" />");
            stringBuilder.AppendLine();
            stringBuilder.Append("<img src=\"");
            stringBuilder.Append(helper.ViewContext.HttpContext.Request.ApplicationPath + "validationImage.ashx?guid=" + image.UniqueId);
            stringBuilder.Append("\" width=\"");
            stringBuilder.Append(width);
            stringBuilder.Append("\" height=\"");
            stringBuilder.Append(height);
            stringBuilder.Append("\"");
            if (!String.IsNullOrWhiteSpace(classValue))
            {
                stringBuilder.Append(String.Format(" class=\"{0}\"", classValue));
            }
            if (!String.IsNullOrWhiteSpace(styleValue))
            {
                stringBuilder.Append(String.Format(" style=\"{0}\"", styleValue));
            }
            stringBuilder.Append(" />");

            return MvcHtmlString.Create(stringBuilder.ToString());
        }
    }
}
