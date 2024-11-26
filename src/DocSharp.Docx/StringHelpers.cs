﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx;

internal static class StringHelpers
{
    internal static string GetLeadingSpaces(string s)
    {
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != ' ')
            {
                return s.Substring(0, i);
            }
        }
        return s;
    }

    internal static string GetTrailingSpaces(string s)
    {
        for (int i = s.Length - 1; i >= 0; i--)
        {
            if (s[i] != ' ')
            {
                if (i < s.Length - 1)
                {
                    return s.Substring(i + 1);
                }
                else
                {
                    return string.Empty;
                }
            }
        }
        return s;
    }
}
