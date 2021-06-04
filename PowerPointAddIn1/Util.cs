using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    class Util
    {
        public static List<PowerPoint.Shape> ListSelectedShapes()
        {
            var shapes = new List<PowerPoint.Shape>();
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return shapes;
            }

            var shapeRange = selection.ShapeRange;

            if (shapeRange == null || shapeRange.Count == 0)
            {
                return shapes;
            }
            foreach (PowerPoint.Shape shape in shapeRange)
            {
                shapes.Add(shape);
            }

            return shapes;
        }

        public static PowerPoint.Shape GetLeftMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }

            PowerPoint.Shape leftMostShape = shapes[0];
            float minLeft = leftMostShape.Left;
            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var left = shape.Left;
                if (left < minLeft)
                {
                    minLeft = left;
                    leftMostShape = shape;
                }
            }
            return leftMostShape;
        }

        public static PowerPoint.Shape GetRightMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }

            PowerPoint.Shape rightMostShape = shapes[0];
            float maxRight = rightMostShape.Left + rightMostShape.Width;
            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var right = shape.Left + shape.Width;
                if (right > maxRight)
                {
                    maxRight = right;
                    rightMostShape = shape;
                }
            }
            return rightMostShape;
        }

        public static PowerPoint.Shape GetTopMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }

            PowerPoint.Shape topMostShape = shapes[0];
            float minTop = topMostShape.Top;
            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var top = shape.Top;
                if (top < minTop)
                {
                    minTop = top;
                    topMostShape = shape;
                }
            }
            return topMostShape;
        }

        public static PowerPoint.Shape GetBottomMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }

            PowerPoint.Shape bottomMostShape = shapes[0];
            float maxBottom = bottomMostShape.Top + bottomMostShape.Height;
            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var bottom = shape.Top + shape.Height;
                if (bottom > maxBottom)
                {
                    maxBottom = bottom;
                    bottomMostShape = shape;
                }
            }
            return bottomMostShape;
        }
    }
}
