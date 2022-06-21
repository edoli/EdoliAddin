using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace EdoliAddIn
{
    public static class ImageTool
    {
        public static void InvertImage()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            var shape = shapes[0];
            if (shape.Type == Core.MsoShapeType.msoPicture)
            {
                FilterImage((imageArray, arraySize, image) =>
                {
                    for (int i = 0; i < arraySize; i++)
                    {
                        imageArray[i] = (byte) (255 - imageArray[i]);
                    }
                    return null;
                });

                var slide = Util.CurrentSlide();
                slide.Shapes.Paste();
            }
        }
        public static void TrimImage()
        {
            var shapes = Util.ListSelectedShapes();
            if (shapes.Count == 0)
            {
                return;
            }

            var shape = shapes[0];
            if (shape.Type == Core.MsoShapeType.msoPicture)
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                selection.Copy();

                var image = Clipboard.GetImage();

                var mStream = new MemoryStream();
                image.Save(mStream, ImageFormat.Bmp);
                var pixelSize = image.Height * image.Width * 4;
                var arraySize = (int)mStream.Length;
                int offset = arraySize - pixelSize;

                var pixelArray = new byte[pixelSize];
                mStream.Position = 54;
                mStream.Read(pixelArray, 0, pixelSize);

                int width = image.Width;
                int height = image.Height;
                var rect = ImageExt.Trim(pixelArray, width, height);

                var shapeWidth = shape.Width;
                var shapeHeight = shape.Height;

                rect = rect.Dilate(-1, width, height);

                shape.PictureFormat.CropLeft += (rect.X * shapeWidth) / width;
                shape.PictureFormat.CropRight += ((width - rect.X - rect.Width) * shapeWidth) / width;
                shape.PictureFormat.CropBottom += (rect.Y * shapeHeight) / height;
                shape.PictureFormat.CropTop += ((height - rect.Y - rect.Height) * shapeHeight) / height;
            }
        }


        public static void FilterImage(Func<byte[], int, Image, byte[]> filter, 
            Func<ImageExt.BmpHeader, ImageExt.BmpHeader> headerFilter = null)
        {
            uint headerSize = 54;
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            selection.Copy();

            var image = Clipboard.GetImage();

            var mStream = new MemoryStream();
            image.Save(mStream, ImageFormat.Bmp);
            var pixelSize = image.Height * image.Width * 4;
            var arraySize = (int) mStream.Length;
            int offset = arraySize - pixelSize;

            var imageArray = mStream.ToArray();
            var pixelArray = new byte[pixelSize];
            Array.Copy(imageArray, offset, pixelArray, 0, pixelSize);

            var headerArray = new byte[headerSize];
            Array.Copy(imageArray, 0, headerArray, 0, headerSize);
            var headerStruct = ImageExt.ByteArrayToStructure<ImageExt.BmpHeader>(headerArray);

            var newArray = filter(pixelArray, pixelSize, image);
            if (newArray != null)
            {
                pixelArray = newArray;
                imageArray = new byte[pixelArray.Length + headerSize];
            }
            if (headerFilter != null)
            {
                headerStruct = headerFilter(headerStruct);
            }
            headerStruct.bfSize = headerStruct.biWidth * headerStruct.biHeight * 4 + headerSize;

            ImageExt.StructureToByteArray(headerStruct).CopyTo(imageArray, 0);
            pixelArray.CopyTo(imageArray, offset);


            var newImage = Image.FromStream(new MemoryStream(imageArray));

            Clipboard.SetImage(newImage);
        }


        public static void ReadImageFromSelection()
        {
            // Globals.ThisAddIn.Application.StartNewUndoEntry();

            var slide = Util.CurrentSlide();
            slide.Shapes.Paste();

        }

        public static void Paste()
        {
            // Globals.ThisAddIn.Application.ActiveWindow.View.PasteSpecial()
        }
    }
}
