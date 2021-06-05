using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    class AnimationTool
    {
        public static void SetNameOfActive(String name)
        {
            var shapes = Util.ListSelectedShapes();

            if (shapes.Count > 0)
            {
                var shape = shapes[0];
                shape.Name = name;
            }
        }

    }
}
