using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class OperatorTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public OperatorTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Equal", 3 },
                { "NotEqual", 4 },
                { "Greater", 5 },
                { "GreaterEqual", 7 },
                { "Less", 6 },
                { "LessEqual", 8 },
                { "Between", 1 },
                { "NotBetween", 2 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new OperatorTypes().Types[name];
        }
    }
}
