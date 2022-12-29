using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Drawing.Imaging;

namespace EdoliAddIn
{
    public static class Util
    {
        public static PowerPoint.Slide CurrentSlide()
        {
            return Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
        }

        public static List<PowerPoint.Shape> ListSlideShapes(PowerPoint.Slide slide = null)
        {
            var shapes = new List<PowerPoint.Shape>();

            if (slide == null)
            {
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            }

            foreach (var shape in slide.Shapes)
            {
                shapes.Add((PowerPoint.Shape) shape);
            }
            return shapes;
        }

        public static List<PowerPoint.Shape> ListSelectedShapes()
        {
            var shapes = new List<PowerPoint.Shape>();
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            var isShapeSelection = (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes
                || selection.Type == PowerPoint.PpSelectionType.ppSelectionText);

            if (selection == null || !isShapeSelection)
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

        public class ImageData
        {
            public byte[] pixelArray;
            public int width;
            public int height;

            public ImageData(byte[] pixelArray, int width, int height)
            {
                this.pixelArray = pixelArray;
                this.width = width;
                this.height = height;
            }

        }

        public static ImageData GetClipboardImageData()
        {
            var image = Clipboard.GetImage();

            var mStream = new MemoryStream();
            image.Save(mStream, ImageFormat.Bmp);
            var pixelSize = image.Height * image.Width * 4;

            var pixelArray = new byte[pixelSize];
            mStream.Position = 54;
            mStream.Read(pixelArray, 0, pixelSize);

            int width = image.Width;
            int height = image.Height;

            Clipboard.Clear();

            return new ImageData(pixelArray, width, height);
        }

        public static bool IsLine(PowerPoint.Shape shape)
        {
            var line = shape.Line;
            return shape.Type == Microsoft.Office.Core.MsoShapeType.msoLine || (int) line.EndArrowheadStyle > 1 || (int) line.BeginArrowheadStyle > 1;
        }


        public class LinePoints
        {
            public float x1;
            public float y1;
            public float x2;
            public float y2;

            public LinePoints() { }

            public LinePoints(float x1, float y1, float x2, float y2)
            {
                this.x1 = x1;
                this.y1 = y1;
                this.x2 = x2;
                this.y2 = y2;
            }
        }

        public static LinePoints GetLinePoints(PowerPoint.Shape shape)
        {
            shape.Copy();

            var imageData = GetClipboardImageData();
            var pixelArray = imageData.pixelArray;
            var width = imageData.width;
            var height = imageData.height;

            var index1 = 0;
            var index2 = (width - 1) * 4;
            var index3 = (height - 1) * width * 4;
            var index4 = (height - 1) * width * 4 + (width - 1) * 4;

            byte a1 = pixelArray[index1 + 3];
            byte a2 = pixelArray[index2 + 3];
            byte a3 = pixelArray[index3 + 3];
            byte a4 = pixelArray[index4 + 3];


            var linePoints = new LinePoints();

            if (a1 > 0 || a4 > 0)
            {
                linePoints.x1 = shape.Left();
                linePoints.y1 = shape.Bottom();
                linePoints.x2 = shape.Right();
                linePoints.y2 = shape.Top();
            } 
            else if (a2 > 0 || a3 > 0)
            {
                linePoints.x1 = shape.Left();
                linePoints.y1 = shape.Top();
                linePoints.x2 = shape.Right();
                linePoints.y2 = shape.Bottom();

            }
            return linePoints;
        }

        public static int NumCluster(IEnumerable<float> values, float threshold)
        {
            var list = values.ToList();
            list.Sort();

            var distances = new float[list.Count - 1];
            var distanceDiffs = new float[list.Count - 2];

            for (int i = 0; i < list.Count() - 1; i++)
            {
                distances[i] = list[i + 1] - list[i];
            }
            Array.Sort(distances);

            //for (int i = 0; i < distances.Count() - 1; i++)
            //{
            //    distanceDiffs[i] = distances[i + 1] - distances[i];
            //}

            //int argmax = Array.IndexOf(distanceDiffs, distanceDiffs.Max());
            //return distances.Count() - argmax;

            for (int i = 0; i < distances.Count(); i++)
            {
                if (distances[i] > threshold)
                {
                    return distances.Count() - i + 1;
                }
            }
            return 1;
        }

        public static TSource MinBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> selector)
        {
            return source.MinBy(selector, null);
        }

        public static TSource MinBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> selector, IComparer<TKey> comparer)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (selector == null) throw new ArgumentNullException("selector");
            comparer = Comparer<TKey>.Default;

            using (var sourceIterator = source.GetEnumerator())
            {
                if (!sourceIterator.MoveNext())
                {
                    throw new InvalidOperationException("Sequence contains no elements");
                }
                var min = sourceIterator.Current;
                var minKey = selector(min);
                while (sourceIterator.MoveNext())
                {
                    var candidate = sourceIterator.Current;
                    var candidateProjected = selector(candidate);
                    if (comparer.Compare(candidateProjected, minKey) < 0)
                    {
                        min = candidate;
                        minKey = candidateProjected;
                    }
                }
                return min;
            }
        }


        public static TSource MaxBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> selector)
        {
            return source.MaxBy(selector, null);
        }

        public static TSource MaxBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> selector, IComparer<TKey> comparer)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (selector == null) throw new ArgumentNullException("selector");
            comparer = Comparer<TKey>.Default;

            using (var sourceIterator = source.GetEnumerator())
            {
                if (!sourceIterator.MoveNext())
                {
                    throw new InvalidOperationException("Sequence contains no elements");
                }
                var max = sourceIterator.Current;
                var maxKey = selector(max);
                while (sourceIterator.MoveNext())
                {
                    var candidate = sourceIterator.Current;
                    var candidateProjected = selector(candidate);
                    if (comparer.Compare(candidateProjected, maxKey) > 0)
                    {
                        max = candidate;
                        maxKey = candidateProjected;
                    }
                }
                return max;
            }
        }

        public static float Dist(float x1, float y1, float x2, float y2)
        {
            float dx = x2 - x1;
            float dy = y2 - y1;
            return (float)Math.Sqrt(dx * dx + dy * dy);
        }

        public static float RectanglePointDistance(float left, float top, float right, float bottom, float x, float y)
        {
            bool isLeft = x < left;
            bool isRight = x > right;
            bool isTop = y < top;
            bool isBottom = y > bottom;

            if (isTop && isLeft)
                return Dist(left, top, x, y);
            else if (isLeft && isBottom)
                return Dist(left, bottom, x, y);
            else if (isBottom && isRight)
                return Dist(right, bottom, x, y);
            else if (isRight && isTop)
                return Dist(right, top, x, y);
            else if (isLeft)
                return left - x;
            else if (isRight)
                return x - right;
            else if (isBottom)
                return y - bottom;
            else if (isTop)
                return top - y;
            else  // contains
                return 0;
        }

        public static float RectangleDistance(float leftA, float topA, float rightA, float bottomA,
            float leftB, float topB, float rightB, float bottomB)
        {
            bool left = rightB < leftA;
            bool right = rightA < leftB;
            bool bottom = bottomB < topA;
            bool top = bottomA < topB;

            if (top && left)
                return Dist(leftA, bottomA, rightB, topB);
            else if (left && bottom)
                return Dist(leftA, topA, rightB, bottomB);
            else if (bottom && right)
                return Dist(rightA, topA, leftB, bottomB);
            else if (right && top)
                return Dist(rightA, bottomA, leftB, topB);
            else if (left)
                return leftA - rightB;
            else if (right)
                return leftB - rightA;
            else if (bottom)
                return topA - bottomB;
            else if (top)
                return topB - bottomA;
            else  // rectangles intersect
                return 0;
        }
    }
}
