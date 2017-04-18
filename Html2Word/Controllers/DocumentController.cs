using Html2Word.Models;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using System.Web.Http.Cors;

namespace Html2Word.Controllers
{
    [EnableCors(origins: "http://localhost:5000", headers: "*", methods: "*")] // TODO: Change origin to the correct URL
    public class DocumentController : ApiController
    {
        /// <summary>
        /// The entrance to the API to get the actual file.
        /// </summary>
        /// <param name="data">The HTML that will be processed.</param>
        /// <returns>Url with the location of the file.</returns>
        [HttpPost]
        public HttpResponseMessage Generate(HtmlPost data)
        {
            try
            {
                // Create a new Word document zip file
                var document = new WordDocument();
                document.CreateDocument(data.Html);

                // Get the url and name of the zip file
                var baseUrl = Request.RequestUri.GetLeftPart(UriPartial.Authority);
                var zipName = document.GUID + ".zip";

                // Create a new response with plain text the url of the zip file
                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(baseUrl + "/Temp/" + zipName, System.Text.Encoding.UTF8, "text/plain")
                };
                return response;
            }
            catch (ArgumentException ex)
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }

        /// <summary>
        /// Get the actual file download.
        /// </summary>
        /// <param name="filePath">The path to save the temp docx.</param>
        /// <returns></returns>
        /// <remarks>This method is not being used, but is saved for future use.</remarks>
        private static HttpResponseMessage GetFile(string filePath)
        {
            var fileName = System.IO.Path.GetFileName(filePath);

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(new FileStream(filePath, FileMode.Open, FileAccess.Read))
            };
            response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = fileName;
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/msword");

            return response;
        }
    }
}
