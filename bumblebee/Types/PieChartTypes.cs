using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace bumblebee.Types
{
    [IsVisibleInDynamoLibrary(false)]
    public class PieChartTypes
    {
        public Dictionary<string, int> Types { get; set; }

        public PieChartTypes()
        {
            Types = new Dictionary<string, int>
            {
                { "3dPie", -4102 },
                { "3dPieExploded", 70 },
                { "Pie", 5 },
                { "PieExploded", 69 }
            };
        }

        public static int ByName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException(nameof(name));
            return new PieChartTypes().Types[name];
        }
    }
}
