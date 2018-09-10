using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class PatternTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public PatternTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "xlCheckerBoard", 9 },
                { "xlCrissCross", 16 },
                { "xlDarkDiagonalDown", -4121 },
                { "xlGrey16", 17 },
                { "xlGray25", -4124 },
                { "xlGray50", -4124 },
                { "xlGray75", -4126 },
                { "xlGray8", 18 },
                { "xlGrid", 15 },
                { "xlDarkHorizontal", -4128 },
                { "xlLightDiagonalDown", 13 },
                { "xlLightHorizontal", 11 },
                { "xlLightDiagonalUp", 14 },
                { "xlLightVertical", 12 },
                { "xlNone", -4142 },
                { "xlSemiGray75", 10 },
                { "xlSolid", 1 },
                { "xlDarkDiagonalUp", -4162 },
                { "xlDarkVertical", -4166 },
                { "mso10Percent", 2 },
                { "mso20Percent", 3 },
                { "mso25Percent", 4 },
                { "msoDarkHorizontal", 13 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new PatternTypes().Types[name];
        }
    }
}
