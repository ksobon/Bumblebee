using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class HorizontalAlignTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public HorizontalAlignTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Left", -4131 },
                { "Center", -4108 },
                { "Right", -4152 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new HorizontalAlignTypes().Types[name];
        }
    }
}
