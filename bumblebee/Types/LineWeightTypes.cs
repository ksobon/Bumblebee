using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class LineWeightTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public LineWeightTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Hairline", 1 },
                { "Medium", -4138 },
                { "Thick", 4 },
                { "Thin", 2 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new LineWeightTypes().Types[name];
        }
    }
}
