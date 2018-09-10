using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class LabelPositionTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public LabelPositionTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Above", 0 },
                { "Below", 1 },
                { "BestFit", 5 },
                { "Center", -4108 },
                { "Custom", 7 },
                { "InsideBase", 4 },
                { "InsideEnd", 3 },
                { "Left", -4131 },
                { "Mixed", 6 },
                { "OutsideEnd", 2 },
                { "Right", -4152 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new LabelPositionTypes().Types[name];
        }
    }
}
