using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCommander.Base
{
    public static class StringHelper
    {
        public static string[] SplitParameters(this string inputString)
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