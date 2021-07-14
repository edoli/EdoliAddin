using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public static class ImageExt
    {
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BmpHeader
        {
            public UInt16 bfType;
            public UInt32 bfSize;
            public UInt16 bfReserved1;
            public UInt16 bfReserved2;
            public UInt32 bfOffBits;

            public UInt32 biSize;
            public UInt32 biWidth;
            public UInt32 biHeight;
            public UInt16 biPlanes;
            public UInt16 biBitCount;
            public UInt32 biCompression;
            public UInt32 biSizeImage;
            public UInt32 biXPelsPerMeter;
            public UInt32 biYPelsPerMeter;
            public UInt32 biClrUsed;
            public UInt32 biClrImportant;
        }

        public static T ByteArrayToStructure<T>(byte[] bytes) where T : struct
        {
            var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                return (T)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(T));
            }
            finally
            {
                handle.Free();
            }
        }

        public static byte[] StructureToByteArray(object obj)
        {
            int len = Marshal.SizeOf(obj);
            byte[] arr = new byte[len];
            IntPtr ptr = Marshal.AllocHGlobal(len);
            Marshal.StructureToPtr(obj, ptr, true);
            Marshal.Copy(ptr, arr, 0, len);
            Marshal.FreeHGlobal(ptr);
            return arr;
        }
        /*
        public static void FixBiImageSize(BmpHeader header)
        {
            uint width = header.biWidth;
            uint height = header.biHeight;
            uint dummy = width % 4;
            dummy = (dummy == 0) ? 0 : 4 - dummy;
            header.biSizeImage = (width + dummy) * height;
        }
        */

        public static float ColorDistance(byte r1, byte g1, byte b1, byte r2, byte g2, byte b2)
        {
            int rd = (int)r1 - r2;
            int gd = (int)g1 - g2;
            int bd = (int)b1 - b2;
            return (float)Math.Sqrt(rd * rd + gd * gd + bd * bd);
        }

        public static Rectangle Trim(byte[] imageArray, int width, int height, float threshold = 10)
        {
            int columnInterval = width * 4;

            int xStart = 0;
            int yStart = 0;
            int xEnd = width;
            int yEnd = height;

            int refIndex = (height - 1) * width * 4;

            byte b = imageArray[refIndex + 0];
            byte g = imageArray[refIndex + 1];
            byte r = imageArray[refIndex + 2];

            // row scan
            for (int i = 0; i < height; i++)
            {
                var index = i * width * 4;
                var cumpin = true;
                for (int j = 0; j < width; j++)
                {
                    var pin = ColorDistance(r, g, b, imageArray[index + 2], imageArray[index + 1], imageArray[index]) < threshold;
                    if (!pin)
                    {
                        cumpin = false;
                        break;
                    }
                    index += 4;
                }

                if (cumpin)
                {
                    yStart = i + 1;
                }
                else
                {
                    break;
                }
            }

            for (int i = height - 1; i >= 0; i--)
            {
                var index = i * width * 4;
                var cumpin = true;
                for (int j = 0; j < width; j++)
                {
                    var pin = ColorDistance(r, g, b, imageArray[index + 2], imageArray[index + 1], imageArray[index]) < threshold;
                    if (!pin)
                    {
                        cumpin = false;
                        break;
                    }
                    index += 4;
                }

                if (cumpin)
                {
                    yEnd = i;
                }
                else
                {
                    break;
                }
            }

            // column scan
            for (int j = 0; j < width; j++)
            {
                var index = j * 4;
                var cumpin = true;
                for (int i = 0; i < height; i++)
                {
                    var pin = ColorDistance(r, g, b, imageArray[index + 2], imageArray[index + 1], imageArray[index]) < threshold;
                    if (!pin)
                    {
                        cumpin = false;
                        break;
                    }
                    index += columnInterval;
                }

                if (cumpin)
                {
                    xStart = j + 1;
                }
                else
                {
                    break;
                }
            }

            for (int j = width - 1; j >= 0; j--)
            {
                var index = j * 4;
                var cumpin = true;
                for (int i = 0; i < height; i++)
                {
                    var pin = ColorDistance(r, g, b, imageArray[index + 2], imageArray[index + 1], imageArray[index]) < threshold;
                    if (!pin)
                    {
                        cumpin = false;
                        break;
                    }
                    index += columnInterval;
                }

                if (cumpin)
                {
                    xEnd = j;
                }
                else
                {
                    break;
                }
            }
            return new Rectangle(xStart, yStart, xEnd - xStart, yEnd - yStart);
        }

        public static Rectangle Dilate(this Rectangle rectangle, int amount = 1, int xBound = -1, int yBound = -1)
        {
            int x = rectangle.X;
            int y = rectangle.Y;
            int width = rectangle.Width;
            int height = rectangle.Height;

            if (x >= amount) {
                x -= amount;
                width += amount;
            }
            if (y >= amount)
            {
                y -= amount;
                height += amount;
            }
            if (xBound == -1 || x + width + amount <= xBound)
            {
                width += amount;
            }
            if (yBound == -1 || y + height + amount <= yBound)
            {
                height += amount;
            }
            return new Rectangle(x, y, width, height);
        }
    }
}
