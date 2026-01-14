using System.Reflection;
using System.Diagnostics;

namespace EPayments.Common.Helpers
{
    public static class AssemblyHelper
    {
        public static string GetAssemblyVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            return fvi.ProductVersion;
        }
    }
}
