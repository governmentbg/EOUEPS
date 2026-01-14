using System.Collections.Generic;
using System.Linq;

namespace EPayments.Common.Linq
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> WithOffsetAndLimit<T>(this IEnumerable<T> query, int offset, int? limit)
        {
            if (offset > 0)
            {
                query = query.Skip(offset);
            }

            if (limit.HasValue)
            {
                query = query.Take(limit.Value);
            }

            return query;
        }
    }
}
