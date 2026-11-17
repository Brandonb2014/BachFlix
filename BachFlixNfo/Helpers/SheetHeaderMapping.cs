using System;
using System.Collections.Generic;

namespace BachFlixNfo.Features
{
    internal static class SheetHeaderMapping
    {
        internal static string ColumnLetter(int column)
        {
            int n = column + 1; string letters = "";
            while (n > 0) { n--; letters = (char)('A' + n % 26) + letters; n /= 26; }
            return letters;
        }
        // Shared with the existing Program.UpdateSheetVariables entry point.
        internal static Dictionary<string, int> Update(Dictionary<string, int> map, IList<IList<object>> headers)
        {
            if (headers != null)
                foreach (var row in headers)
                    for (int i = 0; i < row.Count; i++)
                        if (map.ContainsKey(Convert.ToString(row[i]))) map[Convert.ToString(row[i])] = i;
            return map;
        }
    }
}
