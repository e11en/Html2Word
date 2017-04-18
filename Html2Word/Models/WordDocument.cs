using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using System.Xml.Linq;

namespace Html2Word.Models
{
    public class WordDocument
    {
        public string FilePath { get; set; }

        public WordDocument()
        {
            FilePath = HttpContext.Current.Server.MapPath("~/Temp/temp.html");
        }

        public void CreateDocument(string html)
        {
            html = ProcessImages(html);
            System.IO.File.WriteAllText(FilePath, html);
            System.IO.File.Move(FilePath, HttpContext.Current.Server.MapPath("~/Temp/temp.doc"));
            FilePath = HttpContext.Current.Server.MapPath("~/Temp/temp.doc");
        }

        private static string ProcessImages(string html)
        {
            //Load xml
            XDocument xdoc = XDocument.Parse(html);

            //Run query
            var images = from image in xdoc.Descendants("img")
                         where image.Attribute("src").Value.Contains("base64")
                        select image;
                         

            //Loop through results
            foreach (var img in images)
            {
                var base64 = GetBase64FromDataUri(img.Attribute("src").Value);
                var image = Base64ToImage(base64);
                var imageName = Guid.NewGuid().GetHashCode() + ".png";
                System.IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath("~/Temp/temp_files"));
                image.Save(HttpContext.Current.Server.MapPath("~/Temp/temp_files/") + imageName);
                img.Attribute("src").Value = "./temp_files/" + imageName;
            }

            return xdoc.ToString();
        }

        public static Stream GenerateStreamFromString(string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }


        /// <summary>
        /// Get an Image object from a Base64 image string.
        /// </summary>
        /// <param name="base64String">The string without the "data:image/png;base64," bit.</param>
        /// <returns></returns>
        private static Image Base64ToImage(string base64String)
        {
            // Convert Base64 String to byte[]
            byte[] imageBytes = Convert.FromBase64String(base64String);
            using (var imageStream = new MemoryStream(imageBytes, 0, imageBytes.Length))
            {
                return Image.FromStream(imageStream);
            }
        }

        /// <summary>
        /// Gets the base64 from data URI.
        /// </summary>
        /// <param name="base64String">The base64 string.</param>
        /// <returns></returns>
        private static string GetBase64FromDataUri(string base64String)
        {
            // Remove the 'data:image/png;base64,' bit from an image string.
            return base64String.Split(',')[1];
        }

    }
}
