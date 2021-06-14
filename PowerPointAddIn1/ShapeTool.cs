using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public static class ShapeTool
    {
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
    }
}
