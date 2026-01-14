using System.Web;
using System.Web.Optimization;

namespace EPayments.Web
{
    public class BundleConfig
    {
        // For more information on bundling, visit http://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                "~/Scripts/jquery-1.11.2.min.js",
                "~/Scripts/jquery.easing.1.3.js",
                "~/Scripts/chosen.jquery.min-1.8.7.js"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                "~/Scripts/bootstrap.min.js",
                "~/Scripts/bootstrap-confirmation.js",
                "~/Scripts/bootstrap-datepicker.min.js",
                "~/Scripts/bootstrap-datepicker.bg.min.js",
                "~/Scripts/bootstrap-select.min.js"));

            bundles.Add(new ScriptBundle("~/bundles/upxjs").Include(
                "~/Scripts/upx.js"));

            bundles.Add(new ScriptBundle("~/bundles/homejs").Include(
               "~/Scripts/home.js"));

            bundles.Add(new StyleBundle("~/bundles/css").Include(
                "~/Content/css/fonts.css",
                "~/Content/css/bootstrap.min.css",
                "~/Content/css/bootstrap-datepicker3.standalone.min.css",
                "~/Content/css/style_v3.css",
                "~/Content/css/chosen.min.css"));
        }
    }
}
