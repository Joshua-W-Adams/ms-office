// Refer to the Ghostscript.NET.Samples Projects RasterizerSample.cs file for additional information
// https://github.com/jhabjan/Ghostscript.NET/blob/master/Ghostscript.NET.Samples/

// required Ghostscript.NET namespaces
using System.Drawing.Imaging;
using System.IO;
using Ghostscript.NET.Rasterizer;

namespace Ghostscript.NET.Samples {
    /// <summary>
    /// GhostscriptRasterizer allows you to rasterize pdf and postscript files into the 
    /// memory. If you want Ghostscript to store files on the disk use GhostscriptProcessor
    /// or one of the GhostscriptDevices (GhostscriptPngDevice, GhostscriptJpgDevice).
    /// </summary>
    public class Rasterizer {

        public string pdfToJpeg(string inputPdfFullFileName, string outputPath) {

            // dpi = dots per inch i.e. resolution and quality of jpeg to be created
            // the higher the dpi the longer the processing time.
            int desired_x_dpi = 300;
            int desired_y_dpi = 300;

            string fileName = Path.GetFileNameWithoutExtension(inputPdfFullFileName);
            var outputFullFileName = Path.Combine(outputPath, fileName + ".jpg");

            var rasterizer = new GhostscriptRasterizer();

            using (rasterizer) {

                // open pdf
                rasterizer.Open(inputPdfFullFileName);
                
                // rasterize first page at desired dpi
                var img = rasterizer.GetPage(desired_x_dpi, desired_y_dpi, 1);

                // save image
                img.Save(outputFullFileName, ImageFormat.Jpeg);

            }

            return outputFullFileName;

        }

    }

}