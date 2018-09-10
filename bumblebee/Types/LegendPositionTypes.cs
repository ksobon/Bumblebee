using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class LegendPositionTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public LegendPositionTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Bottom", -4107 },
                { "Upper Right Corner", 2 },
                { "Left", -4131 },
                { "Right", -4152 },
                { "Top", -4160 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new LegendPositionTypes().Types[name];
        }
    }
}
