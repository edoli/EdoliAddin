using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;
using static EdoliAddIn.ShapeExt;
using System.Numerics;

namespace EdoliAddIn
{
    public class AlignTool
    {
        public enum Align
        {
            Top, Bottom, Left, Right
        }

        public static void AlignLeft()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tLeft = lastShape.Left();

                foreach (var shape in shapes)
                {
                    shape.SetLeft(tLeft);
                }
            }
        }
        public static void AlignRight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tRight = lastShape.Right();

                foreach (var shape in shapes)
                {
                    shape.SetRight(tRight);
                }
            }
        }
        public static void AlignTop()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tTop = lastShape.Top();

                foreach (var shape in shapes)
                {
                    shape.SetTop(tTop);
                }
            }
        }
        public static void AlignBottom()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tBottom = lastShape.Bottom();

                foreach (var shape in shapes)
                {
                    shape.SetBottom(tBottom);
                }
            }
        }

        public static void AlignLeftOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tLeft = lastShape.Left();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetRight(tLeft);
                }
            }
        }
        public static void AlignRightOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tRight = lastShape.Right();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetLeft(tRight);
                }
            }
        }
        public static void AlignTopOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tTop = lastShape.Top();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetBottom(tTop);
                }
            }
        }
        public static void AlignBottomOf()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tBottom = lastShape.Bottom();

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.SetTop(tBottom);
                }
            }
        }

        public static void AlignCenter()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

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
            Globals.ThisAddIn.Application.StartNewUndoEntry();

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
            Globals.ThisAddIn.Application.StartNewUndoEntry();

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
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                shapes.Sort((shapeA, shapeB) => Math.Sign(shapeA.Left - shapeB.Left));

                var leftMostShape = ShapeExt.GetLeftMostShape(shapes);
                var rightMostShape = ShapeExt.GetRightMostShape(shapes);

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
        public static void AlignLabels(Anchor anchor)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            var images = new List<PowerPoint.Shape>();
            var textboxes = new List<PowerPoint.Shape>();

            foreach (var shape in shapes)
            {
                if (shape.HasTextFrame == Core.MsoTriState.msoFalse
                    || shape.AutoShapeType == Core.MsoAutoShapeType.msoShapeMixed
                    || shape.TextFrame.TextRange.Text.Equals(""))
                {
                    images.Add(shape);
                }
                else
                {
                    textboxes.Add(shape);
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                }
            }

            foreach (var textbox in textboxes)
            {
                var textAnchor = anchor.Opposite();
                var nearestImage = textbox.FindNearestShape(images, textAnchor, anchor);

                switch (anchor)
                {
                    case Anchor.Top:
                        textbox.SetLeft(nearestImage.Left() + nearestImage.Width() / 2 - textbox.Width() / 2);
                        textbox.SetBottom(nearestImage.Top());
                        break;
                    case Anchor.Bottom:
                        textbox.SetLeft(nearestImage.Left() + nearestImage.Width() / 2 - textbox.Width() / 2);
                        textbox.SetTop(nearestImage.Bottom());
                        break;
                    case Anchor.Left:
                        textbox.SetRight(nearestImage.Left());
                        textbox.SetTop(nearestImage.Top() + nearestImage.Height() / 2 - textbox.Height() / 2);
                        break;
                    case Anchor.Right:
                        textbox.SetLeft(nearestImage.Right());
                        textbox.SetTop(nearestImage.Top() + nearestImage.Height() / 2 - textbox.Height() / 2);
                        break;
                }
            }
        }

        public static void GroupLabels()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            var images = new List<PowerPoint.Shape>();
            var textboxes = new List<PowerPoint.Shape>();

            foreach (var shape in shapes)
            {
                if (shape.HasTextFrame == Core.MsoTriState.msoFalse
                    || shape.AutoShapeType == Core.MsoAutoShapeType.msoShapeMixed
                    || shape.TextFrame.TextRange.Text.Equals(""))
                {
                    images.Add(shape);
                }
                else
                {
                    textboxes.Add(shape);
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                }
            }

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            foreach (var textbox in textboxes)
            {
                try
                {
                    var nearestImage = textbox.FindNearestShape(images, Anchor.None);
                    slide.Shapes.Range(new string[] { textbox.Name, nearestImage.Name }).Group();
                }
                catch
                {

                }
            }
        }

        public static void Transpose()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            var minLeft = shapes.Min(shape => shape.Left());
            var maxLeft = shapes.Max(shape => shape.Left());

            var minTop = shapes.Min(shape => shape.Top());
            var maxTop = shapes.Max(shape => shape.Top());

            var diag = new Vector2(maxLeft - minLeft, maxTop - minTop);

            foreach (var shape in shapes)
            {
                float x = shape.Left() - minLeft;
                float y = shape.Top() - minTop;
                float newX = y * diag.X / diag.Y;
                float newY = x * diag.Y / diag.X;
                shape.SetLeft(newX + minLeft);
                shape.SetTop(newY + minTop);
            }
        }

        public static void AlignGrid()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            float padding = 0;
            int numColumn = 0;
            try
            {
                padding = float.Parse(Globals.Ribbons.EdoliRibbon.gridPadding.Text);
                numColumn = int.Parse(Globals.Ribbons.EdoliRibbon.gridNumColumn.Text);
            }
            catch
            {
                return;
            }

            if (numColumn < 1)
            {
                numColumn = int.MaxValue;
            }

            float left = 0;
            float top = shapes[0].Top();
            float maxHeight = 0;

            for (int i = 0; i < shapes.Count(); i++)
            {
                int col = i % numColumn;
                int row = i / numColumn;
                var shape = shapes[i];

                if (col == 0)
                {
                    left = shapes[0].Left();
                    if (row >= 1)
                    {
                        top += maxHeight + padding;
                    }
                    maxHeight = 0;
                }
                shape.SetLeft(left);
                shape.SetTop(top);

                left += shapes[col].Width() + padding;
                var height = shape.Height();
                if (height > maxHeight)
                {
                    maxHeight = height;
                }
            }
        }

        public static void AlignWithSiblingSlide(int indexOffset)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var slides = Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides;

            int index = slide.SlideIndex;
            int siblingSlideIndex = index + indexOffset;

            if (siblingSlideIndex < 1 || siblingSlideIndex > slides.Count)
            {
                return;
            }

            var prevSlide = slides[siblingSlideIndex];

            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            IEnumerable<PowerPoint.Shape> shapes;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionNone
                || selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                shapes = Util.ListSlideShapes();
            }
            else
            {
                shapes = Util.ListSelectedShapes();
            }

            var prevShapes = Util.ListSlideShapes(prevSlide);

            foreach (var shape in shapes)
            {
                var matchedShape = shape.FindNearestShape(prevShapes, Anchor.Center);
                shape.SetLeft(matchedShape.Left());
                shape.SetTop(matchedShape.Top());

                shape.SetSize(matchedShape.Width(), shape.Height() * matchedShape.Width() / shape.Width());
            }
        }

        public static void SwapCycle()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count < 2)
            {
                return;
            }

            float firstLeft = shapes[0].Left();
            float firstTop = shapes[0].Top();

            for (int i = 0; i < shapes.Count - 1; i++)
            {
                shapes[i].SetLeft(shapes[i + 1].Left());
                shapes[i].SetTop(shapes[i + 1].Top());
            }

            shapes.Last().SetLeft(firstLeft);
            shapes.Last().SetTop(firstTop);
        }

        public static void SwapCycleReverse()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            if (shapes.Count < 2)
            {
                return;
            }

            float lastLeft = shapes.Last().Left();
            float lastTop = shapes.Last().Top();

            for (int i = shapes.Count - 1; i > 0; i--)
            {
                shapes[i].SetLeft(shapes[i - 1].Left());
                shapes[i].SetTop(shapes[i - 1].Top());
            }
            shapes[0].SetLeft(lastLeft);
            shapes[0].SetTop(lastTop);
        }

        public static void SnapDownRight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            var firstShape = shapes[0];
            var left = firstShape.Right();
            var top = firstShape.Bottom();

            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                shape.SetLeft(left);
                shape.SetTop(top);

                top += shape.Height();
                left += shape.Width();
            }

        }

        public static void SnapUpRight()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            var firstShape = shapes[0];
            var left = firstShape.Right();
            var top = firstShape.Top();

            for (int i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                top -= shape.Height();
                shape.SetLeft(left);
                shape.SetTop(top);
                left += shape.Width();
            }
        }

        public static void AlignHorizontalVertical()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();

            if (shapes.Count > 0)
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                float minLeft = shapes.Min(s => s.Left);
                float minTop = shapes.Min(s => s.Top);

                float maxLeft = shapes.Max(s => s.Left);
                float maxTop = shapes.Max(s => s.Top);

                float meanWidth = shapes.Average(s => s.Width);
                float meanHeight = shapes.Average(s => s.Height);

                var lefts = shapes.Select(s => s.Left);
                var tops = shapes.Select(s => s.Top);

                int numHorizontalCluster = Util.NumCluster(lefts, meanWidth);
                int numVerticalCluster = Util.NumCluster(tops, meanHeight);

                float hInterval = 0;
                float vInterval = 0;

                if (numHorizontalCluster > 1)
                {
                    hInterval = (maxLeft - minLeft) / (numHorizontalCluster - 1);
                }

                if (numVerticalCluster > 1)
                {
                    vInterval = (maxTop - minTop) / (numVerticalCluster - 1);
                }

                foreach (var shape in shapes)
                {
                    int row = (int)((shape.Top - minTop + vInterval / 2) / vInterval);
                    shape.Top = minTop + row * vInterval;

                    int col = (int)((shape.Left - minLeft + hInterval / 2) / hInterval);
                    shape.Left = minLeft + col * hInterval;
                }
            }
        }

        public static void BringToFront()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoBringToFront);
            });
        }

        public static void SendToBack()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoSendToBack);
            });
        }

        public static void BringForward()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoBringForward);
            });
        }

        public static void SendBackward()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var shapes = Util.ListSelectedShapes();
            shapes.ForEach(shape =>
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoSendBackward);
            });
        }
    }
}
