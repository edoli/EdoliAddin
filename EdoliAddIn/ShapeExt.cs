using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public static class ShapeExt
    {
        public enum Anchor
        {
            Top, Bottom, Left, Right,
            TopLeft, TopRight, BottomLeft, BottomRight,
            Center,
            None
        }

        public static Anchor Opposite(this Anchor anchor)
        {
            switch (anchor)
            {
                case Anchor.Top: return Anchor.Bottom;
                case Anchor.Bottom: return Anchor.Top;
                case Anchor.Left: return Anchor.Right;
                case Anchor.Right: return Anchor.Left;
                case Anchor.TopLeft: return Anchor.BottomRight;
                case Anchor.TopRight: return Anchor.BottomLeft;
                case Anchor.BottomLeft: return Anchor.TopRight;
                case Anchor.BottomRight: return Anchor.TopLeft;
                case Anchor.Center: return Anchor.Center;
                case Anchor.None: return Anchor.None;
                default: return Anchor.None;

            }
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

        public static float Width(this PowerPoint.Shape shape)
        {
            float rotation = (float)(shape.Rotation * Math.PI / 180.0f);
            return (float)(Math.Abs(Math.Cos(rotation)) * shape.Width + Math.Abs(Math.Sin(rotation)) * shape.Height);
        }

        public static float Height(this PowerPoint.Shape shape)
        {
            float rotation = (float)(shape.Rotation * Math.PI / 180.0f);
            return (float)(Math.Abs(Math.Sin(rotation)) * shape.Width + Math.Abs(Math.Cos(rotation)) * shape.Height);
        }

        public static void SetSize(this PowerPoint.Shape shape, float width, float height)
        {
            float rotation = (float)(shape.Rotation * Math.PI / 180.0f);
            float cps = (float)(Math.Abs(Math.Cos(rotation)) + Math.Abs(Math.Sin(rotation)));
            float cms = (float)(Math.Abs(Math.Cos(rotation)) - Math.Abs(Math.Sin(rotation)));
            float wph = (width + height) / cps;
            float wmh = (width - height) / cms;
            shape.Width = (wph + wmh) / 2.0f;
            shape.Height = (wph - wmh) / 2.0f;
        }

        public static float Left(this PowerPoint.Shape shape)
        {
            float offset = (Width(shape) - shape.Width) / 2;
            return shape.Left - offset;
        }

        public static void SetLeft(this PowerPoint.Shape shape, float value)
        {
            float offset = (Width(shape) - shape.Width) / 2;
            shape.Left = value + offset;
        }

        public static float Right(this PowerPoint.Shape shape)
        {
            float width = Width(shape);
            float offset = (width - shape.Width) / 2;
            return shape.Left + width - offset;
        }

        public static void SetRight(this PowerPoint.Shape shape, float value)
        {
            float width = Width(shape);
            float offset = (width - shape.Width) / 2;
            shape.Left = value - width + offset;
        }

        public static float Top(this PowerPoint.Shape shape)
        {
            float offset = (Height(shape) - shape.Height) / 2;
            return shape.Top - offset;
        }

        public static void SetTop(this PowerPoint.Shape shape, float value)
        {
            float offset = (Height(shape) - shape.Height) / 2;
            shape.Top = value + offset;
        }

        public static float Bottom(this PowerPoint.Shape shape)
        {
            float height = Height(shape);
            float offset = (height - shape.Height) / 2;
            return shape.Top + height - offset;
        }

        public static void SetBottom(this PowerPoint.Shape shape, float value)
        {
            float height = Height(shape);
            float offset = (height - shape.Height) / 2;
            shape.Top = value - height + offset;
        }

        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB)
        {
            var left1 = shapeA.Left();
            var right1 = shapeA.Right();
            var top1 = shapeA.Top();
            var bottom1 = shapeA.Bottom();

            var left2 = shapeB.Left();
            var right2 = shapeB.Right();
            var top2 = shapeB.Top();
            var bottom2 = shapeB.Bottom();

            return Util.RectangleDistance(left1, top1, right1, bottom1, left2, top2, right2, bottom2);
        }

        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB, Anchor anchor)
        {
            if (anchor == Anchor.None)
            {
                return DistanceOfShapes(shapeA, shapeB);
            } 
            else
            {
                var p1 = shapeA.Position(anchor);
                var p2 = shapeB.Position(anchor);
                return Vector2.Distance(p1, p2);
            }
        }
        public static float DistanceOfShapes(PowerPoint.Shape shapeA, PowerPoint.Shape shapeB, Anchor anchorA, Anchor anchorB)
        {
            if (anchorA == Anchor.None && anchorB == Anchor.None)
            {
                return DistanceOfShapes(shapeA, shapeB);
            }
            else if (anchorA == Anchor.None)
            {
                var p2 = shapeB.Position(anchorB);
                return Util.RectanglePointDistance(shapeA.Left(), shapeA.Top(), shapeA.Right(), shapeA.Bottom(), p2.X, p2.Y);
            }
            else if (anchorB == Anchor.None)
            {
                var p1 = shapeA.Position(anchorA);
                return Util.RectanglePointDistance(shapeB.Left(), shapeB.Top(), shapeB.Right(), shapeB.Bottom(), p1.X, p1.Y);
            }
            else
            {
                var p1 = shapeA.Position(anchorA);
                var p2 = shapeB.Position(anchorB);
                return Vector2.Distance(p1, p2);
            }
        }

        public static Vector2 Position(this PowerPoint.Shape shape, Anchor anchor)
        {
            switch (anchor)
            {
                case Anchor.Left:
                    return new Vector2(shape.Left(), shape.Top() + shape.Height() / 2);
                case Anchor.Right:
                    return new Vector2(shape.Left() + shape.Width(), shape.Top() + shape.Height() / 2);
                case Anchor.Top:
                    return new Vector2(shape.Left() + shape.Width() / 2, shape.Top());
                case Anchor.Bottom:
                    return new Vector2(shape.Left() + shape.Width() / 2, shape.Top() + shape.Height());
                case Anchor.TopLeft:
                    return new Vector2(shape.Left(), shape.Top());
                case Anchor.TopRight:
                    return new Vector2(shape.Left() + shape.Width(), shape.Top());
                case Anchor.BottomLeft:
                    return new Vector2(shape.Left(), shape.Top() + shape.Height());
                case Anchor.BottomRight:
                    return new Vector2(shape.Left() + shape.Width(), shape.Top() + shape.Height());
                case Anchor.Center:
                    return new Vector2(shape.Left() + shape.Width() / 2, shape.Top() + shape.Height() / 2);
            }
            return new Vector2();
        }

        public static PowerPoint.Shape FindNearestShape(this PowerPoint.Shape shape, List<PowerPoint.Shape> shapes, Anchor anchor)
        {
            return shapes.MinBy(s => DistanceOfShapes(shape, s, anchor));
        }

        public static PowerPoint.Shape FindNearestShape(this PowerPoint.Shape shape, List<PowerPoint.Shape> shapes, Anchor anchorA, Anchor anchorB)
        {
            return shapes.MinBy(s => DistanceOfShapes(shape, s, anchorA, anchorB));
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

        public static Anchor GetRelativePos(this PowerPoint.Shape shape, PowerPoint.Shape other, float epsilon=0)
        {
            float left1 = shape.Left();
            float top1 = shape.Top();
            float right1 = shape.Right();
            float bottom1 = shape.Bottom();

            float left2 = other.Left();
            float top2 = other.Top();
            float right2 = other.Right();
            float bottom2 = other.Bottom();

            if (right2 < left1 - epsilon && bottom2 < top1 - epsilon)
            {
                return Anchor.TopLeft;
            }
            if (left2 > right1 + epsilon && bottom2 < top1 - epsilon)
            {
                return Anchor.TopRight;
            }
            if (right2 < left1 - epsilon && top2 > bottom1 + epsilon)
            {
                return Anchor.BottomLeft;
            }
            if (left2 > right1 + epsilon && top2 > bottom1 + epsilon)
            {
                return Anchor.BottomRight;
            }
            if (right2 < left1)
            {
                return Anchor.Left;
            }
            if (left2 > right1)
            {
                return Anchor.Right;
            }
            if (bottom2 < top1)
            {
                return Anchor.Top;
            }
            if (top2 > bottom1)
            {
                return Anchor.Bottom;
            }
            return Anchor.None;
        }
    }
}
