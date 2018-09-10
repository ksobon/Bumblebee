using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class LineChartTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public LineChartTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "Line", 4 },
                { "LineStacked", 63 },
                { "LineStacked100", 64 },
                { "3dLine", -4101 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new LineChartTypes().Types[name];
        }
    }
}
