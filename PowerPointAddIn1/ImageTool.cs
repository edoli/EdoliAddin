using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public static class ImageTool
    {
        public static void InvertImage()
        {
            var shapes = Util.ListSelectedShapes();

            foreach(var shape in shapes)
            {
                if (shape.Type == Core.MsoShapeType.msoPicture)
                {
                    
                }
            }
            // Home End 단축키
            // 이미지 필터 단축키 (brightness)
        }
    }
}
