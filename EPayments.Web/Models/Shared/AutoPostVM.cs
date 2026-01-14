using System.Collections.Generic;

namespace EPayments.Web.Models.Shared
{
    public class AutoPostVM
    {

        public string PostUrl { get; set; }
        public List<KeyValuePair<string, string>> PostValues { get; set; }
        public string MethodType { get; set; } = "post";
    }
}