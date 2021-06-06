using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public static class ShapeTool
    {
        public static void ChangeStroke(float offset)
        {
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                shape.DoRecur(s =>
                {
                    var line = shape.Line;
                    if (line.Style > 0)
                    {
                        if (line.Weight > -offset)
                        {
                            line.Weight += offset;
                        }
                    }
                });
            }
        }
    }
}
