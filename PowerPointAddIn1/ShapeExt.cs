using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public static class ShapeExt
    {
        public enum Anchor
        {
            Top, Bottom, Left, Right,
            TopLeft, TopRight, BottomLeft, BottomRight,
            None
        }

        public static void DoRecur(this PowerPoint.Shape shape, Action<PowerPoint.Shape> action)
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape item in shape.GroupItems)
                {
                    item.DoRecur(action);
                }
            }
            else
            {
                action(shape);
            }
        }

        public static float Right(this PowerPoint.Shape shape)
        {
            return shape.Left + shape.Width;
        }

        public static void SetRight(this PowerPoint.Shape shape, float right)
        {
            shape.Left = right - shape.Width;
        }
        public static float Bottom(this PowerPoint.Shape shape)
        {
            return shape.Top + shape.Height;
        }

        public static void SetBottom(this PowerPoint.Shape shape, float bottom)
        {
            shape.Top = bottom - shape.Height;
        }


        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB, Anchor anchor)
        {
            if (anchor == Anchor.None)
            {
                var left1 = shapeA.Left;
                var right1 = shapeA.Right();
                var top1 = shapeA.Top;
                var bottom1 = shapeA.Bottom();

                var left2 = shapeB.Left;
                var right2 = shapeB.Right();
                var top2 = shapeB.Top;
                var bottom2 = shapeB.Bottom();

                return Util.RectangleDistance(left1, top1, right1, bottom1, left2, top2, right2, bottom2);
            } 
            else
            {
                var p1 = shapeA.Position(anchor);
                var p2 = shapeB.Position(anchor);
                return Vector2.Distance(p1, p2);
            }
        }

        public static Vector2 Position(this PowerPoint.Shape shape, Anchor anchor)
        {
            switch (anchor)
            {
                case Anchor.Left:
                    return new Vector2(shape.Left, shape.Top + shape.Height / 2);
                case Anchor.Right:
                    return new Vector2(shape.Left + shape.Width, shape.Top + shape.Height / 2);
                case Anchor.Top:
                    return new Vector2(shape.Left + shape.Width / 2, shape.Top);
                case Anchor.Bottom:
                    return new Vector2(shape.Left + shape.Width / 2, shape.Top + shape.Height);
                case Anchor.TopLeft:
                    return new Vector2(shape.Left, shape.Top);
                case Anchor.TopRight:
                    return new Vector2(shape.Left + shape.Width, shape.Top);
                case Anchor.BottomLeft:
                    return new Vector2(shape.Left, shape.Top + shape.Height);
                case Anchor.BottomRight:
                    return new Vector2(shape.Left + shape.Width, shape.Top + shape.Height);
            }
            return new Vector2();
        }

        public static PowerPoint.Shape FindNearestShape(this PowerPoint.Shape shape, List<PowerPoint.Shape> shapes, Anchor anchor)
        {
            return shapes.MinBy(s => DistanceOfShapes(shape, s, anchor));
        }

        public static PowerPoint.Shape GetLeftMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MinBy(shape => shape.Left);
        }

        public static PowerPoint.Shape GetRightMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MaxBy(shape => shape.Right());
        }

        public static PowerPoint.Shape GetTopMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MinBy(shape => shape.Top);
        }

        public static PowerPoint.Shape GetBottomMostShape(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count == 0)
            {
                return null;
            }
            return shapes.MaxBy(shape => shape.Bottom());
        }
    }
}
