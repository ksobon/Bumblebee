using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class ColorScaleCriteriaTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public ColorScaleCriteriaTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "None", -1 },
                { "Number", 0 },
                { "LowestValue", 1 },
                { "HighestValue", 2 },
                { "Percent", 3 },
                { "Formula", 4 },
                { "Percentile", 5 },
                { "AutomaticMin", 6 },
                { "AutomaticMax", 7 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new ColorScaleCriteriaTypes().Types[name];
        }
    }
}
