using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class DirectionTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public DirectionTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "LeftToRight", -5003 },
                { "RightToLeft", -5004 },
                { "Context", -5002 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new DirectionTypes().Types[name];
        }
    }
}
