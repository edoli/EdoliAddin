using System;
using Core = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace EdoliAddIn
{
    public static class ImageTool
    {

        public static void TrimImage()
        {
            
            ActionProcessPowerPointImages((pixelArray, arraySize, image, shape) =>
            {
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
                return false;
            });
        }

        public static void ActionProcessPowerPointImages(
            Func<byte[], int, Image, Shape, bool> action = null,
            Func<ImageExt.BmpHeader, ImageExt.BmpHeader> headerFilter = null)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapes = Util.ListSelectedShapes();

            foreach (var shape in shapes)
            {
                if (shape.Type == Core.MsoShapeType.msoPicture)
                {
                    ProcessPowerPointImage(shape, action, headerFilter);
                }
            }
        }
        

        public static void ProcessPowerPointImage(
            Shape shape,
            Func<byte[], int, Image, Shape, bool> action = null,
            Func<ImageExt.BmpHeader, ImageExt.BmpHeader> headerFilter = null)
        {
            // 이미지를 임시 파일로 내보내기
            string tempPath = Path.GetTempFileName();
            shape.Export(tempPath, PpShapeFormat.ppShapeFormatPNG);

            // 이미지를 임시 파일로 내보내기
            uint headerSize = 54;
            byte[] imageArray;
            bool reloadImage = false;

            using (Bitmap image = new Bitmap(tempPath))
            {
                var pixelSize = image.Height * image.Width * 4;
                var arraySize = pixelSize + headerSize;

                using (MemoryStream ms = new MemoryStream())
                {
                    image.Save(ms, ImageFormat.Bmp);
                    imageArray = ms.ToArray();
                }

                var pixelArray = new byte[pixelSize];
                Array.Copy(imageArray, headerSize, pixelArray, 0, pixelSize);

                var headerArray = new byte[headerSize];
                Array.Copy(imageArray, 0, headerArray, 0, headerSize);
                var headerStruct = ImageExt.ByteArrayToStructure<ImageExt.BmpHeader>(headerArray);

                reloadImage = action?.Invoke(pixelArray, pixelSize, image, shape) ?? false;

                if (reloadImage)
                {
                    Array.Copy(pixelArray, 0, imageArray, headerSize, pixelSize);
                }

                if (headerFilter != null)
                {
                    headerStruct = headerFilter(headerStruct);
                    headerStruct.bfSize = headerStruct.biWidth * headerStruct.biHeight * 4 + headerSize;
                    ImageExt.StructureToByteArray(headerStruct).CopyTo(imageArray, 0);
                }
            }

            // tempPath 으로 불러온 Bitmap 리소스가 해제된 다음에 수행
            if (reloadImage)
            {
                // 처리된 이미지를 다시 임시 파일로 저장
                using (var processedBmp = new Bitmap(new MemoryStream(imageArray)))
                {
                    using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                    {
                        processedBmp.Save(stream, ImageFormat.Png);
                    }
                }

                // 처리된 이미지로 Shape 업데이트
                shape.Fill.UserPicture(tempPath);
            }

            // 임시 파일 삭제
            File.Delete(tempPath);
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
