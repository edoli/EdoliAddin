using Expressive;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Numerics;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public class PointsShape
    {
        public float[,] points;
        public float width;
        public float height;

        public PointsShape(Vector2[] localPoints, float offsetX, float offsetY)
        {
            int numPoints = localPoints.Length;
            points = new float[numPoints, 2];

            var minV = new Vector2(float.MaxValue, float.MaxValue);
            var maxV = new Vector2(float.MinValue, float.MinValue);
            for (int i = 0; i < numPoints; i++)
            {
                var v = localPoints[i];

                if (v.X < minV.X) { minV.X = v.X; }
                if (v.Y < minV.Y) { minV.Y = v.Y; }
                if (v.X > maxV.X) { maxV.X = v.X; }
                if (v.Y > maxV.Y) { maxV.Y = v.Y; }
            }
            float cx = (minV.X + maxV.X) / 2;
            float cy = (minV.Y + maxV.Y) / 2;

            width = maxV.X - minV.X;
            height = maxV.Y - minV.Y;

            for (int i = 0; i < numPoints; i++)
            {
                var v = localPoints[i];
                points[i, 0] = v.X + offsetX - cx;
                points[i, 1] = -v.Y + offsetY + cy;
            }
        }
    }

    public static class ShapeTool
    {

        private static float shapeScale = 28.3465f;

        public static string PathTagName = "EquationPath";
        public static string CurveTag = "EquationCurve";
        public static string PolylineTag = "EquationPolyline";


        public static void ToggleLine()
        {
            var shapes = Util.ListSelectedShapes();

            shapes.ForEach(s =>
            {
                if (s.Line.Visible == MsoTriState.msoFalse)
                {
                    s.Line.Visible = MsoTriState.msoTrue;
                }
                else
                {
                    s.Line.Visible = MsoTriState.msoFalse;
                }
            });
        }

        public static void ChangeLineWeight(float offset)
        {
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                shape.DoRecur(s =>
                {
                    var line = s.Line;
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
        public static void ChangeLineDash(int offset)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                shape.DoRecur(s =>
                {
                    var line = s.Line;
                    if (line.Style > 0)
                    {
                        int style = (int)line.DashStyle;
                        if (style == 2 || style == 3)
                        {
                            style = 1;
                        }
                        if (style > 2)
                        {
                            style -= 2;
                        }

                        int newDashStyle = offset + style;
                        if (newDashStyle > 10)
                        {
                            newDashStyle = 1;
                        }
                        if (newDashStyle < 1)
                        {
                            newDashStyle = 10;
                        }

                        // ignore 2, 3 style
                        if (newDashStyle >= 2)
                        {
                            newDashStyle += 2;
                        }
                        line.DashStyle = (MsoLineDashStyle)newDashStyle;
                    }
                });
            }
        }

        public static void BeginArrowToggle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                try
                {
                    if (shape.Line.BeginArrowheadStyle == MsoArrowheadStyle.msoArrowheadNone)
                    {
                        shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    }
                    else
                    {
                        shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
                    }
                }
                catch
                {

                }
            }
        }

        public static void BeginArrowChangeSize(int deltaSize)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                {
                    if (shape.Line.BeginArrowheadStyle != MsoArrowheadStyle.msoArrowheadNone)
                    {
                        var width = (int)shape.Line.BeginArrowheadWidth;
                        var length = (int)shape.Line.BeginArrowheadLength;

                        var newWidth = width + deltaSize;
                        var newLength = length + deltaSize;

                        if (newWidth > 0 && newWidth <= 3)
                        {
                            shape.Line.BeginArrowheadWidth = (MsoArrowheadWidth)newWidth;
                        }

                        if (newLength > 0 && newLength <= 3)
                        {
                            shape.Line.BeginArrowheadLength = (MsoArrowheadLength)newLength;
                        }
                    }
                }
            }
        }

        public static void EndArrowToggle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                try
                {
                    if (shape.Line.EndArrowheadStyle == MsoArrowheadStyle.msoArrowheadNone)
                    {
                        shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                    }
                    else
                    {
                        shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
                    }
                }
                catch
                {

                }
            }
        }

        public static void EndArrowChangeSize(int deltaSize)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                {
                    if (shape.Line.EndArrowheadStyle != MsoArrowheadStyle.msoArrowheadNone)
                    {
                        var width = (int)shape.Line.EndArrowheadWidth;
                        var length = (int)shape.Line.EndArrowheadLength;

                        var newWidth = width + deltaSize;
                        var newLength = length + deltaSize;

                        if (newWidth > 0 && newWidth <= 3)
                        {
                            shape.Line.EndArrowheadWidth = (MsoArrowheadWidth)newWidth;
                        }

                        if (newLength > 0 && newLength <= 3)
                        {
                            shape.Line.EndArrowheadLength = (MsoArrowheadLength)newLength;
                        }
                    }
                }
            }
        }

        public static void ConnectShapesByLine()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            for (int i = 0; i < shapes.Count - 1; i++)
            {
                var shape1 = shapes[i];
                var shape2 = shapes[i + 1];

                var rel = shape1.GetRelativePos(shape2, 0.1f);

                if (rel == ShapeExt.Anchor.None)
                {
                    continue;
                }

                PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                float left1 = shape1.Left;
                float top1 = shape1.Top;
                float right1 = shape1.Right();
                float bottom1 = shape1.Bottom();

                float left2 = shape2.Left;
                float top2 = shape2.Top;
                float right2 = shape2.Right();
                float bottom2 = shape2.Bottom();

                if (rel == ShapeExt.Anchor.TopLeft || rel == ShapeExt.Anchor.BottomRight)
                {
                    slide.Shapes.AddLine(left1, bottom1, left2, bottom2);
                    slide.Shapes.AddLine(right1, top1, right2, top2);
                }

                if (rel == ShapeExt.Anchor.TopRight || rel == ShapeExt.Anchor.BottomLeft)
                {
                    slide.Shapes.AddLine(right1, bottom1, right2, bottom2);
                    slide.Shapes.AddLine(left1, top1, left2, top2);
                }

                if (rel == ShapeExt.Anchor.Left)
                {
                    slide.Shapes.AddLine(left1, top1, right2, top2);
                    slide.Shapes.AddLine(left1, bottom1, right2, bottom2);
                }

                if (rel == ShapeExt.Anchor.Right)
                {
                    slide.Shapes.AddLine(right1, top1, left2, top2);
                    slide.Shapes.AddLine(right1, bottom1, left2, bottom2);
                }

                if (rel == ShapeExt.Anchor.Top)
                {
                    slide.Shapes.AddLine(left1, top1, left2, bottom2);
                    slide.Shapes.AddLine(right1, top1, right2, bottom2);
                }

                if (rel == ShapeExt.Anchor.Bottom)
                {
                    slide.Shapes.AddLine(left1, bottom1, left2, top2);
                    slide.Shapes.AddLine(right1, bottom1, right2, top2);
                }
            }
        }

        public static void AddPathOfExpression(string expX, string expY, string startValue, string endValue, bool isCurve)
        {
            float startValueEvaluated = Convert.ToSingle(new Expression(startValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            float endValueEvaluated = Convert.ToSingle(new Expression(endValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());

            var expressiveX = new Expression(expX, ExpressiveOptions.IgnoreCaseForParsing);
            var expressiveY = new Expression(expY, ExpressiveOptions.IgnoreCaseForParsing);
            AddPathOfFunction(t => {
                var dict = new Dictionary<string, object> { ["t"] = t };
                return new Vector2(Convert.ToSingle(expressiveX.Evaluate(dict)),
                                   Convert.ToSingle(expressiveY.Evaluate(dict)));
            }, startValueEvaluated, endValueEvaluated, isCurve);
        }

        public static void AddPathOfFunction(Func<float, Vector2> func, float startValue, float endValue, bool isCurve)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var slide = Util.CurrentSlide();
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;

            float slideHeight = currentPresentation.PageSetup.SlideHeight;
            float slideWidth = currentPresentation.PageSetup.SlideWidth;

            try
            {
                var vectors = PathOfFunction(func, startValue, endValue, isCurve);
                var pointsShape = new PointsShape(vectors, slideWidth / 2, slideHeight / 2);
                var points = pointsShape.points;
                var shape = isCurve ? slide.Shapes.AddCurve(points) : slide.Shapes.AddPolyline(points);
                shape.Tags.Add(PathTagName, isCurve ? CurveTag : PolylineTag);
                shape.Select();

                if (Globals.Ribbons.EdoliRibbon.checkBoxNormalizeEqShape.Checked)
                {
                    float scale = 64 / pointsShape.width;
                    shape.ScaleWidth(scale, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);
                    shape.ScaleHeight(scale, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public static void UpdatePathOfExpression(string expX, string expY, string startValue, string endValue)
        {
            // Disable for now

            // float startValueEvaluated = Convert.ToSingle(new Expression(startValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            // float endValueEvaluated = Convert.ToSingle(new Expression(endValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());

            // var expressiveX = new Expression(expX, ExpressiveOptions.IgnoreCaseForParsing);
            // var expressiveY = new Expression(expY, ExpressiveOptions.IgnoreCaseForParsing);
            // UpdatePathOfFunction(t => {
            //     var dict = new Dictionary<string, object> { ["t"] = t };
            //     return new Vector2(Convert.ToSingle(expressiveX.Evaluate(dict)),
            //                        Convert.ToSingle(expressiveY.Evaluate(dict)));
            // }, startValueEvaluated, endValueEvaluated);
        }

        public static void UpdatePathOfFunction(Func<float, Vector2> func, float startValue, float endValue)
        {
            // Globals.ThisAddIn.Application.StartNewUndoEntry();
            var slide = Util.CurrentSlide();
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;
            var selectedShapes = Util.ListSelectedShapes();
            if (selectedShapes.Count == 1)
            {
                var shape = selectedShapes[0];
                var tag = shape.Tags[PathTagName];
                if (tag == CurveTag || tag == PolylineTag)
                {
                    try
                    {
                        bool isCurve = tag == CurveTag;
                        var vectors = PathOfFunction(func, startValue, endValue, isCurve);
                        var center = shape.Position(ShapeExt.Anchor.Center);
                        var pointsShape = new PointsShape(vectors, center.X, center.Y);
                        var points = pointsShape.points;
                        
                        int nodeCount = shape.Nodes.Count;
                        for (int i = 0; i < nodeCount; i++)
                        {
                            shape.Nodes.SetPosition(i + 1, points[i, 0], points[i, 1]);
                        }
                        // shape.Delete();
                        
                        // var newShape = isCurve ? slide.Shapes.AddCurve(points) : slide.Shapes.AddPolyline(points);
                        // newShape.Tags.Add(PathTagName, isCurve ? CurveTag : PolylineTag);
                        // newShape.Select();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        public static Vector2[] PathOfFunction(Func<float, Vector2> func, float startValue, float endValue, bool addControlPoints = false)
        {
            float rangeValue = endValue - startValue;
            var numPoints = 100;
            var initVectors = new Vector2[numPoints];
            for (int t = 0; t < numPoints; t++)
            {
                var f = ((float)t) / (numPoints - 1);
                initVectors[t] = func(startValue + f * rangeValue) * shapeScale;
            }
            
            Vector2[] vectors;
            if (addControlPoints)
            {
                // Add control points
                vectors = new Vector2[numPoints];
                for (int t = 0; t < numPoints; t++)
                {
                    if (t % 3 == 0 || t == 1 || t == numPoints - 2)
                    {
                        vectors[t] = initVectors[t];
                        continue;
                    }

                    Vector2 v1 = initVectors[t];
                    Vector2 v2 = new Vector2();
                    Vector2 v0 = new Vector2();
                    if (t % 3 == 1)
                    {
                        v2 = initVectors[t - 2];
                        v0 = initVectors[t - 1];
                    }
                    if (t % 3 == 2)
                    {
                        v2 = initVectors[t + 2];
                        v0 = initVectors[t + 1];
                    }
                    vectors[t] = v0 + (v1 - v2) / 2;
                }
            }
            else
            {
                vectors = initVectors;
            }

            return vectors;
        }

        public static void DrawAngleLines()
        {

        }

        public static void DrawPerpendicularAngleLines()
        {

        }
    }
}
