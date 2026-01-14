using System.Collections.Generic;

namespace EPayments.Admin.Models.Shared
{
    public class AutoPostVM
    {
        public string PostUrl { get; set; }
        public List<KeyValuePair<string, string>> PostValues { get; set; }
    }
}