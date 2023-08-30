using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSC_Legend_Renderer.Services
{
    class ImageProcessing
    {

        ///https://www.codeguru.com/columns/dotnet/creating-images-from-scratch-in-c.html
        /// <summary>
        /// Will create a mono colored copy of input image with given color object.
        /// The new image will be saved at given path.
        /// </summary>
        /// <param name="imageToMimic">The original image to copy outline from</param>
        /// <param name="inColor">Wanted color for the new image</param>
        /// <param name="outPath">The output path for the new image</param>
        public Bitmap CreateMonoColorFromImageCopy(Image imageToMimic, Color inColor, string outPath, int demTransparency = 178)
        {

            //Calculate and create array buffer size
            int byteBuffer = 4 * imageToMimic.Width * imageToMimic.Height;
            byte[] imageBuffer = new byte[byteBuffer];

            for (int y = 0; y < imageToMimic.Height; y++)
            {
                for (int x = 0; x < imageToMimic.Width; x++)
                {
                    byte red = (byte)inColor.R;
                    byte green = (byte)inColor.G;
                    byte blue = (byte)inColor.B;

                    PlotPixel(x, y, red, green, blue, imageBuffer, imageToMimic, demTransparency);
                }
            }

            unsafe
            {
                fixed (byte* ptr = imageBuffer)
                {
                    using (Bitmap outImage = new Bitmap(imageToMimic.Width, imageToMimic.Height, imageToMimic.Width * 4, PixelFormat.Format32bppArgb, new IntPtr(ptr)))
                    {
                        outImage.Save(outPath);

                        return outImage;
                    }
                }

            }
        }

        /// <summary>
        /// Will plot the color of a pixel in a given byte array
        /// </summary>
        /// <param name="x">x coordinate of the pixel</param>
        /// <param name="y">y coordinate of the pixel</param>
        /// <param name="redValue">The red component for the new pixel</param>
        /// <param name="greenValue">The green component for the new pixel</param>
        /// <param name="blueValue">The blue component for the new pixel</param>
        /// <param name="inBuffer">The byte array that act as buffer</param>
        /// <param name="inImage">The original image to retrieve width and height</param>
        public void PlotPixel(int x, int y, byte redValue, byte greenValue, byte blueValue, byte[] inBuffer, Image inImage, int demTransparency)
        {
            //Calculate dem transparency
            byte transparency = (byte)demTransparency;
            int offset = ((inImage.Width * 4) * y) + (x * 4);
            inBuffer[offset] = blueValue;
            inBuffer[offset + 1] = greenValue;
            inBuffer[offset + 2] = redValue;
            inBuffer[offset + 3] = transparency;
        }

    }
}
