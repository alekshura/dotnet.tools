using System;
using System.Drawing;
using System.IO;

namespace Dotnet.Tools.Barcode
{
    /// <summary>
    /// Interleaved 25 barcode image generator helper class.
    /// See more: https://www.activebarcode.com/codes/code25interleaved.html
    /// </summary>
    public static class c
    {
        public static Image GetBarcodeImage(string barcodeNumber)
        {
            string[] patterns = { "NNWWN",//0 
                                  "WNNNW",//1 
                                  "NWNNW",//2 
                                  "WWNNN",//3 
                                  "NNWNW",//4 
                                  "WNWNN",//5 
                                  "NWWNN",//6 
                                  "NNNWW",//7 
                                  "WNNWN",//8 
                                  "NWNWN" //9 
                                };

            int width = 400;
            int height = 60;

            string barcodeString = "NnNn";//start pattern
            int barcodeLength = barcodeNumber.Length;

            for (int i = 0; i < barcodeLength; i++)
            {
                int firstNumber = int.Parse(barcodeNumber[i].ToString());
                int secondNumber = int.Parse(barcodeNumber[i + 1].ToString());

                string firstPattern = patterns[firstNumber].ToUpper();
                string secondPattern = patterns[secondNumber].ToLower();

                barcodeString += @$"{firstPattern[0]}{secondPattern[0]}{firstPattern[1]}
                    {secondPattern[1]}{firstPattern[2]}{secondPattern[2]}
                    {firstPattern[3]}{secondPattern[3]}{firstPattern[4]}{secondPattern[4]}";

                i++;
            }

            barcodeString += "WnN";//stop pattern

            var barcodeImage = new Bitmap(width, height) as Image;
            using (var gr = Graphics.FromImage(barcodeImage))
            {
                // set to white so we don't have to fill the spaces with white
                gr.FillRectangle(Brushes.White, 0, 0, width, height);

                int cursor = 0;

                for (int codeidx = 0; codeidx < barcodeString.Length; codeidx++)
                {
                    char code = barcodeString[codeidx];

                    int BarWeight = 1;
                    int barwidth = ((code == 'N') || (code == 'n')) ? 2 * BarWeight : 6 * BarWeight;

                    if ((code == 'N') || (code == 'W'))
                    {
                        gr.FillRectangle(Brushes.Black, cursor, 0, barwidth, height);
                    }

                    // note that we never need to draw the space, since we 
                    // initialized the graphics to all white

                    // advance cursor beyond this pair
                    cursor += barwidth;
                }
            }

            return barcodeImage;
        }

        public static byte[] ImageToByteArray(Image image)
        {
            var ms = new MemoryStream();
            image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            return ms.ToArray();
        }

        public static Image StringToImage(string inputString)
        {
            byte[] imageToByte = Convert.FromBase64String(inputString);
            var ms = new MemoryStream(imageToByte);
            Image NewImage = Image.FromStream(ms);
            return NewImage;
        }
    }
}
