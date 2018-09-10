using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class CompoundLineTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public CompoundLineTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "msoSingle", 1 },
                { "msoThinThin", 2 },
                { "msoThinThick", 3 },
                { "msoThickThin", 4 },
                { "msoThickBetweenThin", 5 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new CompoundLineTypes().Types[name];
        }
    }
}
