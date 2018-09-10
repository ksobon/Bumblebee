using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class LineTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public LineTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "xlContinuous", 1 },
                { "xlDash", -4115 },
                { "xlDashDot", 4 },
                { "xlDashDotDot", 5 },
                { "xlRoundDot", -4118 },
                { "xlLongDash", -4115 },
                { "xlDouble", -4119 },
                { "xlNone", -4142 },
                { "msoDash", 4 },
                { "msoDashDot", 5 },
                { "msoDashDotDot", 6 },
                { "msoDashStyleMixed", -2 },
                { "msoLongDash", 7 },
                { "msoLongDashDot", 8 },
                { "msoRoundDot", 3 },
                { "msoSolid", 1 },
                { "msoSquareDot", 2 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new LineTypes().Types[name];
        }
    }
}
