using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class VerticalAlignTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public VerticalAlignTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Bottom", -4017 },
                { "Center", -4108 },
                { "Top", -4160 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new VerticalAlignTypes().Types[name];
        }
    }
}
