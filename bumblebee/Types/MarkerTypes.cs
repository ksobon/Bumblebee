using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class MarkerTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public MarkerTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "xlAutomatic", -4105 },
                { "xlCircle", 8 },
                { "xlDash", -4115 },
                { "xlDiamond", 2 },
                { "xlDot", -4118 },
                { "xlNone", -4142 },
                { "xlPlus", 9 },
                { "xlSquare", 1 },
                { "xlStar", 5 },
                { "xlTriangle", 3 },
                { "xlX", -4168 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new MarkerTypes().Types[name];
        }
    }
}
