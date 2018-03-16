using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GPSTeachingSys
{
    class Example_01
    {
        public static string getPath(string path)
        {
            int t;
            for (t = 0; t < path.Length; t++)
            {
                if (path.Substring(t, 14) == "GPSTeachingSys")
                {
                    break;
                }
            }
            string name = path.Substring(0, t - 1);
            return name;
        }
    }
}
