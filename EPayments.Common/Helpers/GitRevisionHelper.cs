using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Common.Helpers
{
    public class GitRevisionHelper
    {
        private static string _revision = null;
        private static string _revisionFilePath = null;

        private static string RevisionFilePath
        {
            get
            {
                if (_revisionFilePath == null)
                {
                    _revisionFilePath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "GitRevision.txt");
                }

                return _revisionFilePath;
            }
        }

        public static string GetRevision()
        {
            if (_revision == null)
            {
                try
                {
                    using (var reader = new StreamReader(RevisionFilePath))
                    {
                        _revision = reader.ReadLine();
                    }
                }
                catch
                {
                }
            }

            return _revision ?? String.Empty;
        }
    }
}
