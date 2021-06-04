using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public class AlignTool
    {
        public static void AlignLeft()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tLeft = lastShape.Left;

                foreach (var shape in shapes)
                {
                    shape.Left = tLeft;
                }
            }
        }
        public static void AlignRight()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tRight = lastShape.Left + lastShape.Width;

                foreach (var shape in shapes)
                {
                    shape.Left = tRight - shape.Width;
                }
            }
        }
        public static void AlignTop()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tTop = lastShape.Top;

                foreach (var shape in shapes)
                {
                    shape.Top = tTop;
                }
            }
        }
        public static void AlignBottom()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tBottom = lastShape.Top + lastShape.Height;

                foreach (var shape in shapes)
                {
                    shape.Top = tBottom - shape.Height;
                }
            }
        }
        public static void AlignCenter()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float hCenter = lastShape.Left + lastShape.Width / 2;
                float vCenter = lastShape.Top + lastShape.Height / 2;

                foreach (var shape in shapes)
                {
                    shape.Left = hCenter - shape.Width / 2;
                    shape.Top = vCenter - shape.Height / 2;
                }
            }
        }
        public static void AlignCenterHorizontal()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float hCenter = lastShape.Left + lastShape.Width / 2;

                foreach (var shape in shapes)
                {
                    shape.Left = hCenter - shape.Width / 2;
                }
            }
        }
        public static void AlignCenterVertical()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float vCenter = lastShape.Top + lastShape.Height / 2;

                foreach (var shape in shapes)
                {
                    shape.Top = vCenter - shape.Height / 2;
                }
            }
        }

        public static void AlignInRow()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            { 
                shapes.Sort((shapeA, shapeB) => Math.Sign(shapeA.Left - shapeB.Left));

                var leftMostShape = Util.GetLeftMostShape(shapes);
                var rightMostShape = Util.GetRightMostShape(shapes);

                var left = leftMostShape.Left;
                var right = rightMostShape.Left + rightMostShape.Width;
                var top = leftMostShape.Top;

                var height = leftMostShape.Height;
                foreach (var shape in shapes)
                {
                    shape.Width = shape.Width * height / shape.Height;
                    shape.Height = height;
                }

                float sumWidth = 0;
                foreach (var shape in shapes)
                {
                    sumWidth += shape.Width;
                }
                float interval = (right - left - sumWidth) / (shapes.Count - 1);
                float culLeft = left;
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    shape.Left = culLeft;
                    shape.Top = top;

                    culLeft += shape.Width + interval;
                }
            }
        }
    }
}
