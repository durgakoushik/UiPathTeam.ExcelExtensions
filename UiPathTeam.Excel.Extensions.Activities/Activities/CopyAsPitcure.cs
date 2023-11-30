using Microsoft.Office.Interop.Excel;
using System;
using System.Activities;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Media.Imaging;

namespace UiPathTeam.Excel.Extensions.Activities
{
    [DisplayName("CopyAsPitcure")]
    public class CopyAsPitcure : CodeActivity
    {

        [Category("Input")]
        [Description("Path to which you want to save the image.")]
        [RequiredArgument]
        public InArgument<string> Path { get; set; }

        [Category("Input")]
        [Description("If quality of an image needs to be reduced by 50%")]
        [RequiredArgument]
        public bool ReduceQuality { get; set; }
        public CopyAsPitcure()
        {
            ReduceQuality = true;
            Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<CopyAsPitcure>(typeof(ExcelExtensionScope).Name));
        }
        [STAThread()]
        protected override void Execute(CodeActivityContext context)
        {
            var property = context.DataContext.GetProperties()[ExcelExtensionScope.ExcelTag];
            var excelProperty = property.GetValue(context.DataContext) as ExcelSession;

            Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)excelProperty.worksheet.Application.Selection;
            rng.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
            //Image image = null;
            Thread.Sleep(1000);
            if (Clipboard.ContainsImage())
            {
                SaveImage(Path.Get(context), ReduceQuality);
                //var bitmapSource = Clipboard.GetImage();
                //image = new Image(bitmapSource);

                //image = Clipboard.GetImage();

                //if (ReduceQuality)
                //    SaveJpeg(Path.Get(context), image, 50);
                //else
                //    image.Save(Path.Get(context));
            }
            else
            {
                Console.WriteLine("Image not present in clipboard");
            }
        }
        public static void SaveImage(string path, bool ReduceQuality)
        {
            var image = Clipboard.GetImage();
            
            BitmapEncoder encoder = null;
            if (ReduceQuality)
            {
                encoder = new JpegBitmapEncoder();
                ((JpegBitmapEncoder)encoder).QualityLevel = 50;
            }
            else
            {
                encoder = new PngBitmapEncoder();
            }
            encoder.Frames.Add(BitmapFrame.Create(image));

            using (var fileStream = new FileStream(path, FileMode.Create))
            {
                encoder.Save(fileStream);
            }
        }

        //public static void SaveJpeg(string path, Image img, int quality)
        //{
        //    if (quality < 0 || quality > 100)
        //        throw new ArgumentOutOfRangeException("quality must be between 0 and 100.");

        //    // Encoder parameter for image quality 
        //    EncoderParameter qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
        //    // JPEG image codec 
        //    ImageCodecInfo jpegCodec = GetEncoderInfo("image/jpeg");
        //    EncoderParameters encoderParams = new EncoderParameters(1);
        //    encoderParams.Param[0] = qualityParam;
        //    img.Save(path, jpegCodec, encoderParams);
        //}

        /// <summary> 
        /// Returns the image codec with the given mime type 
        /// </summary> 
        //private static ImageCodecInfo GetEncoderInfo(string mimeType)
        //{
        //    // Get image codecs for all image formats 
        //    ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();

        //    // Find the correct image codec 
        //    for (int i = 0; i < codecs.Length; i++)
        //        if (codecs[i].MimeType == mimeType)
        //            return codecs[i];

        //    return null;
        //}
    }
}
