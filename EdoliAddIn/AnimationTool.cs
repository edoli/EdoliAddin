using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    class AnimationTool
    {
        public static void SetNameOfActive(String name)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            if (shapes.Count > 0)
            {
                var shape = shapes[0];
                shape.Name = name;
            }
        }

    }
}
