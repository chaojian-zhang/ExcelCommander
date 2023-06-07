using System.Collections.Generic;
using System.Text;

namespace ExcelCommander.Base
{
    public static class StringHelper
    {
        public static string[] SplitParameters(this string inputString, bool includeQuotesInString = false)
        {
            List<string> parameters = new List<string>();
            StringBuilder current = new StringBuilder();

            bool inQuotes = false;
            foreach (var c in inputString)
            {
                switch (c)
                {
                    case '"':
                        inQuotes = !inQuotes;
                        if (includeQuotesInString)
                            current.Append(c);
                        break;
                    case ' ':
                        if (!inQuotes)
                        {
                            parameters.Add(current.ToString());
                            current.Clear();
                        }
                        else
                            current.Append(c);
                        break;
                    default:
                        current.Append(c);
                        break;
                }
            }
            if (current.Length != 0)
                parameters.Add(current.ToString());
            return parameters.ToArray();
        }
    }
}