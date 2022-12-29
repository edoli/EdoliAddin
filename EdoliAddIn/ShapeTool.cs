using Expressive;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace EdoliAddIn
{
    public static class ShapeTool
    {

        private static float shapeScale = 28.3465f;
        public static void ToggleLine()
        {
            var shapes = Util.ListSelectedShapes();

            shapes.ForEach(s =>
            {
                if (s.Line.Visible == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    s.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                }
                else
                {
                    s.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
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
                        line.DashStyle = (Microsoft.Office.Core.MsoLineDashStyle)newDashStyle;
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
                    if (shape.Line.BeginArrowheadStyle == Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone)
                    {
                        shape.Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle;
                    }
                    else
                    {
                        shape.Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone;
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
                    if (shape.Line.BeginArrowheadStyle != Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone)
                    {
                        var width = (int)shape.Line.BeginArrowheadWidth;
                        var length = (int)shape.Line.BeginArrowheadLength;

                        var newWidth = width + deltaSize;
                        var newLength = length + deltaSize;

                        if (newWidth > 0 && newWidth <= 3)
                        {
                            shape.Line.BeginArrowheadWidth = (Microsoft.Office.Core.MsoArrowheadWidth)newWidth;
                        }

                        if (newLength > 0 && newLength <= 3)
                        {
                            shape.Line.BeginArrowheadLength = (Microsoft.Office.Core.MsoArrowheadLength)newLength;
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
                    if (shape.Line.EndArrowheadStyle == Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone)
                    {
                        shape.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle;
                    }
                    else
                    {
                        shape.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone;
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
                    if (shape.Line.EndArrowheadStyle != Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone)
                    {
                        var width = (int)shape.Line.EndArrowheadWidth;
                        var length = (int)shape.Line.EndArrowheadLength;

                        var newWidth = width + deltaSize;
                        var newLength = length + deltaSize;

                        if (newWidth > 0 && newWidth <= 3)
                        {
                            shape.Line.EndArrowheadWidth = (Microsoft.Office.Core.MsoArrowheadWidth)newWidth;
                        }

                        if (newLength > 0 && newLength <= 3)
                        {
                            shape.Line.EndArrowheadLength = (Microsoft.Office.Core.MsoArrowheadLength)newLength;
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

        public static void AddCurveOfExpression(string expX, string expY, string startValue, string endValue)
        {
            float startValueEvaluated = Convert.ToSingle(new Expression(startValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            float endValueEvaluated = Convert.ToSingle(new Expression(endValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());

            var expressiveX = new Expression(expX, ExpressiveOptions.IgnoreCaseForParsing);
            var expressiveY = new Expression(expY, ExpressiveOptions.IgnoreCaseForParsing);
            AddCurveOfFunction(t => {
                var dict = new Dictionary<string, object> { ["t"] = t };
                return new Vector2(Convert.ToSingle(expressiveX.Evaluate(dict)),
                                   Convert.ToSingle(expressiveY.Evaluate(dict)));
            }, startValueEvaluated, endValueEvaluated);
        }

        public static void AddCurveOfFunction(Func<float, Vector2> func, float startValue, float endValue)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var slide = Util.CurrentSlide();
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;

            float slideHeight = currentPresentation.PageSetup.SlideHeight;
            float slideWidth = currentPresentation.PageSetup.SlideWidth;

            float rangeValue = endValue - startValue;

            try
            {
                var numPoints = 100;
                var initVectors = new Vector2[numPoints];
                for (int t = 0; t < numPoints; t++)
                {
                    var f = ((float)t) / (numPoints - 1);
                    initVectors[t] = func(startValue + f * rangeValue) * shapeScale;
                }

                var vectors = new Vector2[numPoints];
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

                var points = new float[numPoints, 2];

                var minV = new Vector2(float.MaxValue, float.MaxValue);
                var maxV = new Vector2(float.MinValue, float.MinValue);
                for (int i = 0; i < numPoints; i++)
                {
                    var v = vectors[i];

                    if (v.X < minV.X) { minV.X = v.X; }
                    if (v.Y < minV.Y) { minV.Y = v.Y; }
                    if (v.X > maxV.X) { maxV.X = v.X; }
                    if (v.Y > maxV.Y) { maxV.Y = v.Y; }
                }
                float cx = (minV.X + maxV.X) / 2;
                float cy = (minV.Y + maxV.Y) / 2;

                float width = maxV.X - minV.X;
                float height = maxV.Y - minV.Y;

                for (int i = 0; i < numPoints; i++)
                {
                    var v = vectors[i];
                    points[i, 0] = v.X + slideWidth / 2 - cx;
                    points[i, 1] = -v.Y + slideHeight / 2 + cy;
                }

                var shape = slide.Shapes.AddCurve(points);

                if (Globals.Ribbons.EdoliRibbon.checkBoxNormalizeEqShape.Checked)
                {
                    float scale = 64 / width;
                    shape.ScaleWidth(scale, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                    shape.ScaleHeight(scale, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public static void AddPolylineOfExpression(string expX, string expY, string startValue, string endValue)
        {
            float startValueEvaluated = Convert.ToSingle(new Expression(startValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            float endValueEvaluated = Convert.ToSingle(new Expression(endValue, ExpressiveOptions.IgnoreCaseForParsing).Evaluate());
            var expressiveX = new Expression(expX, ExpressiveOptions.IgnoreCaseForParsing);
            var expressiveY = new Expression(expY, ExpressiveOptions.IgnoreCaseForParsing);
            AddPolylineOfFunction(t => {
                var dict = new Dictionary<string, object> { ["t"] = t };
                return new Vector2(Convert.ToSingle(expressiveX.Evaluate(dict)),
                                   Convert.ToSingle(expressiveY.Evaluate(dict)));
            }, startValueEvaluated, endValueEvaluated);
        }


        public static void AddPolylineOfFunction(Func<float, Vector2> func, float startValue, float endValue)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var slide = Util.CurrentSlide();
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;

            float slideHeight = currentPresentation.PageSetup.SlideHeight;
            float slideWidth = currentPresentation.PageSetup.SlideWidth;

            float rangeValue = endValue - startValue;

            try
            {
                var numPoints = 100;
                var vectors = new Vector2[numPoints];
                for (int t = 0; t < numPoints; t++)
                {
                    var f = ((float)t) / (numPoints - 1);
                    vectors[t] = func(startValue + f * rangeValue) * shapeScale;
                }

                var points = new float[numPoints, 2];

                var minV = new Vector2(float.MaxValue, float.MaxValue);
                var maxV = new Vector2(float.MinValue, float.MinValue);
                for (int i = 0; i < numPoints; i++)
                {
                    var v = vectors[i];

                    if (v.X < minV.X) { minV.X = v.X; }
                    if (v.Y < minV.Y) { minV.Y = v.Y; }
                    if (v.X > maxV.X) { maxV.X = v.X; }
                    if (v.Y > maxV.Y) { maxV.Y = v.Y; }
                }
                float cx = (minV.X + maxV.X) / 2;
                float cy = (minV.Y + maxV.Y) / 2;

                float width = maxV.X - minV.X;
                float height = maxV.Y - minV.Y;

                for (int i = 0; i < numPoints; i++)
                {
                    var v = vectors[i];
                    points[i, 0] = v.X + slideWidth / 2 - cx;
                    points[i, 1] = -v.Y + slideHeight / 2 + cy;
                }

                var shape = slide.Shapes.AddPolyline(points);
                if (Globals.Ribbons.EdoliRibbon.checkBoxNormalizeEqShape.Checked)
                {
                    float scale = 64 / width;
                    shape.ScaleWidth(scale, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                    shape.ScaleHeight(scale, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public static void AlignLines()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();

                var lastLinePoints = Util.GetLinePoints(lastShape);

                foreach (var shape in shapes)
                {
                    if (Util.IsLine(shape))
                    {
                        Util.GetLinePoints(shape);
                    }
                }
            }
        }

        public static void TrimLines()
        {

        }

        public static void DrawAngleLines()
        {

        }

        public static void DrawPerpendicularAngleLines()
        {

        }
    }
}
