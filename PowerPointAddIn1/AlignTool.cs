using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;
using static PowerPointAddIn1.ShapeExt;
using System.Numerics;

namespace PowerPointAddIn1
{
    public class AlignTool
    {
        public enum Align
        {
            Top, Bottom, Left, Right
        }

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

        public static void AlignLeftOf()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tLeft = lastShape.Left;

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.Left = tLeft - shape.Width;
                }
            }
        }
        public static void AlignRightOf()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tRight = lastShape.Left + lastShape.Width;

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.Left = tRight;
                }
            }
        }
        public static void AlignTopOf()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tTop = lastShape.Top;

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.Top = tTop - shape.Height;
                }
            }
        }
        public static void AlignBottomOf()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count > 1)
            {
                var lastShape = shapes.Last();
                float tBottom = lastShape.Top + lastShape.Height;

                for (int i = 0; i < shapes.Count - 1; i++)
                {
                    var shape = shapes[i];
                    shape.Top = tBottom;
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
                var nearestImage = textbox.FindNearestShape(images, Anchor.None);

                switch (anchor)
                {
                    case Anchor.Top:
                        textbox.Left = nearestImage.Left + nearestImage.Width / 2 - textbox.Width / 2;
                        textbox.SetBottom(nearestImage.Top);
                        break;
                    case Anchor.Bottom:
                        textbox.Left = nearestImage.Left + nearestImage.Width / 2 - textbox.Width / 2;
                        textbox.Top = nearestImage.Bottom();
                        break;
                    case Anchor.Left:
                        textbox.SetRight(nearestImage.Left);
                        textbox.Top = nearestImage.Top + nearestImage.Height / 2 - textbox.Height / 2;
                        break;
                    case Anchor.Right:
                        textbox.Left = nearestImage.Right();
                        textbox.Top = nearestImage.Top + nearestImage.Height / 2 - textbox.Height / 2;
                        break;
                }
            }
        }

        public static void GroupLabels()
        {
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
            var shapes = Util.ListSelectedShapes();

            var minLeft = shapes.Min(shape => shape.Left);
            var maxLeft = shapes.Max(shape => shape.Left);

            var minTop = shapes.Min(shape => shape.Top);
            var maxTop = shapes.Max(shape => shape.Top);

            var diag = new Vector2(maxLeft - minLeft, maxTop - minTop);

            foreach (var shape in shapes)
            {
                float x = shape.Left - minLeft;
                float y = shape.Top - minTop;
                float newX = y * diag.X / diag.Y;
                float newY = x * diag.Y / diag.X;
                shape.Left = newX + minLeft;
                shape.Top = newY + minTop;
            }
        }

        public static void AlignGrid()
        {
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
            float top = shapes[0].Top;
            float maxHeight = 0;

            for (int i = 0; i < shapes.Count(); i++)
            {
                int col = i % numColumn;
                int row = i / numColumn;
                var shape = shapes[i];

                if (col == 0)
                {
                    left = shapes[0].Left;
                    if (row >= 1)
                    {
                        top += maxHeight + padding;
                    }
                    maxHeight = 0;
                }
                shape.Left = left;
                shape.Top = top;
                left += shapes[col].Width + padding;
                if (shape.Height > maxHeight)
                {
                    maxHeight = shape.Height;
                }
            }
        }

        public static void AlignWithPreviousSlide()
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var index = slide.SlideIndex;

            if (index <= 1)
            {
                return;
            }

            var prevSlide = Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[index - 1];

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
                var matchedShape = shape.FindNearestShape(prevShapes, Anchor.TopLeft);
                shape.Left = matchedShape.Left;
                shape.Top = matchedShape.Top;

                shape.Height = shape.Height * matchedShape.Width / shape.Width;
                shape.Width = matchedShape.Width;
            }
        }

        public static void SwapCycle()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count < 2)
            {
                return;
            }

            float firstLeft = shapes[0].Left;
            float firstTop = shapes[0].Top;

            for (int i = 0; i < shapes.Count - 1; i++)
            {
                shapes[i].Left = shapes[i + 1].Left;
                shapes[i].Top = shapes[i + 1].Top;
            }

            shapes.Last().Left = firstLeft;
            shapes.Last().Top = firstTop;
        }

        public static void SwapCycleReverse()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count < 2)
            {
                return;
            }

            float lastLeft = shapes.Last().Left;
            float lastTop = shapes.Last().Top;

            for (int i = shapes.Count - 1; i > 0; i--)
            {
                shapes[i].Left = shapes[i - 1].Left;
                shapes[i].Top = shapes[i - 1].Top;
            }
            shapes[0].Left = lastLeft;
            shapes[0].Top = lastTop;
        }

        public static void SnapUpLeft()
        {
            var shapes = Util.ListSelectedShapes();



            for (int i = 1; i < shapes.Count; i++)
            {
                shapes[i].Left = shapes[i + 1].Left;
                shapes[i].Top = shapes[i + 1].Top;
            }

        }
    }
}
