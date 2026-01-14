using System.Web;
using System.Web.Optimization;

namespace EPayments.Admin
{
    public class BundleConfig
    {
        // For more information on bundling, visit https://go.microsoft.com/fwlink/?LinkId=301862
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
                "~/Scripts/bootstrap-datepicker.bg.min.js"));

            bundles.Add(new ScriptBundle("~/bundles/upxjs").Include(
                "~/Scripts/upx.js",
                "~/Scripts/clipboard.min.js"));

            bundles.Add(new ScriptBundle("~/bundles/homejs").Include(
               "~/Scripts/adminhome.js"));

            bundles.Add(new StyleBundle("~/bundles/css").Include(
                "~/Content/css/fonts.css",
                "~/Content/css/bootstrap.min.css",
                "~/Content/css/bootstrap-datepicker3.standalone.min.css",
                "~/Content/css/style_v2.css",
                "~/Content/css/chosen.min.css"));
        }
    }
}
