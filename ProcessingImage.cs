using System.Drawing;
using System.Drawing.Imaging;

public class ProcessingImage
{

    public static void MakeTransparentImage(string inputPath, string outputPath, float opacity)
    {
        using (var image = Image.FromFile(inputPath))
        using (var bitmap = new Bitmap(image.Width, image.Height))
        using (var g = Graphics.FromImage(bitmap))
        {
            var matrix = new ColorMatrix
            {
                Matrix33 = opacity // 0.0 = fully transparent, 1.0 = original
            };

            var attributes = new ImageAttributes();
            attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

            g.DrawImage(
                image,
                new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                0, 0, image.Width, image.Height,
                GraphicsUnit.Pixel,
                attributes);

            bitmap.Save(outputPath, ImageFormat.Png);
        }
    }

}