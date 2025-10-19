using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;

class ConvertSvgToIco
{
    static void Main()
    {
        try
        {
            // Read SVG file
            string svgContent = File.ReadAllText("ui\\logo.svg");

            // Extract base64 data
            Match match = Regex.Match(svgContent, @"xlink:href=""data:img/png;base64,([^""]+)""");
            if (!match.Success)
            {
                Console.WriteLine("Error: Cannot extract base64 data from SVG");
                Environment.Exit(1);
            }

            string base64Data = match.Groups[1].Value;
            byte[] imageBytes = Convert.FromBase64String(base64Data);

            // Load image
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Image sourceImage = Image.FromStream(ms))
            {
                // Create ICO file with multiple sizes
                int[] sizes = { 16, 32, 48, 64, 128, 256 };
                using (MemoryStream iconStream = new MemoryStream())
                using (BinaryWriter writer = new BinaryWriter(iconStream))
                {
                    // Write ICO header
                    writer.Write((ushort)0);  // Reserved
                    writer.Write((ushort)1);  // Type (1 = ICO)
                    writer.Write((ushort)sizes.Length);  // Number of images

                    // Prepare image data
                    MemoryStream[] imageStreams = new MemoryStream[sizes.Length];
                    for (int i = 0; i < sizes.Length; i++)
                    {
                        int size = sizes[i];
                        using (Bitmap bitmap = new Bitmap(size, size))
                        {
                            using (Graphics g = Graphics.FromImage(bitmap))
                            {
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.DrawImage(sourceImage, 0, 0, size, size);
                            }

                            imageStreams[i] = new MemoryStream();
                            bitmap.Save(imageStreams[i], ImageFormat.Png);
                        }
                    }

                    // Write directory entries
                    long offset = 6 + (16 * sizes.Length);
                    for (int i = 0; i < sizes.Length; i++)
                    {
                        int size = sizes[i];
                        byte[] data = imageStreams[i].ToArray();

                        writer.Write((byte)(size == 256 ? 0 : size));  // Width
                        writer.Write((byte)(size == 256 ? 0 : size));  // Height
                        writer.Write((byte)0);   // Color palette
                        writer.Write((byte)0);   // Reserved
                        writer.Write((ushort)1); // Color planes
                        writer.Write((ushort)32); // Bits per pixel
                        writer.Write((uint)data.Length); // Image data size
                        writer.Write((uint)offset); // Image data offset

                        offset += data.Length;
                    }

                    // Write image data
                    for (int i = 0; i < imageStreams.Length; i++)
                    {
                        byte[] data = imageStreams[i].ToArray();
                        writer.Write(data);
                        imageStreams[i].Dispose();
                    }

                    // Save ICO file
                    File.WriteAllBytes("icon.ico", iconStream.ToArray());
                }
            }

            Console.WriteLine("Successfully created icon.ico!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            Environment.Exit(1);
        }
    }
}
