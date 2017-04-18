using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml.Linq;
using System.IO.Compression;

namespace Html2Word.Models
{
    public class WordDocument
    {
        public string FilePath { get; set; }
        public string GUID { get; set; }

        public WordDocument()
        {
            GUID = Guid.NewGuid().GetHashCode().ToString();
            FilePath = HttpContext.Current.Server.MapPath("~/Temp/" + GUID + "/");
        }

        public void CreateDocument(string html)
        {
            // Save the images to a temp_files folder and replace image src
            html = ProcessImages(html);

            // Write the doc file
            System.IO.File.WriteAllText(FilePath + "temp.doc", html);

            // Create a zip file of the doc and image folder
            ZipFile.CreateFromDirectory(FilePath, HttpContext.Current.Server.MapPath("~/Temp/" + GUID + ".zip"));

            // Remove folder
            Directory.Delete(FilePath, true);
        }

        /// <summary>
        /// Extract the Base64 images from the HTML and save them in a temp_files folder
        /// then point the src to the saved image.
        /// </summary>
        /// <param name="html">The complete HTML.</param>
        /// <returns></returns>
        private string ProcessImages(string html)
        {
            // Load the html as xml
            var xdoc = XDocument.Parse(html);

            // Get all the images
            var images = from image in xdoc.Descendants("img")
                         where image.Attribute("src").Value.Contains("base64")
                        select image;
                         

            foreach (var img in images)
            {
                // Get Image object from image src
                var base64 = GetBase64FromDataUri(img.Attribute("src").Value);
                var image = Base64ToImage(base64);

                // Save image to the temp_files folder
                var imageName = Guid.NewGuid().GetHashCode() + ".png";
                System.IO.Directory.CreateDirectory(FilePath + "temp_files");
                image.Save(FilePath + "temp_files/" + imageName);

                // Rename the src to the saved image filename
                img.Attribute("src").Value = "./temp_files/" + imageName;
            }

            // Return the complete html string with new image names
            return xdoc.ToString();
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
