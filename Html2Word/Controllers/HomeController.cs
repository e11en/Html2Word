using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web.Http;

namespace Html2Word.Controllers
{
    public class HomeController : ApiController
    {
        [HttpGet]
        public HttpResponseMessage Index()
        {
            var response = new HttpResponseMessage
            {
                Content = new StringContent(
                    "<html><body><h1>Document Zip Generate</h1>Use in Angular example: <br> <table border='1'><tr><td>" +
                    "$http.post('" + Request.RequestUri.GetLeftPart(UriPartial.Authority) + "/api/document/generate',{ Html : html });" +
                    "</td></tr></table></body></html>"
                    )
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            return response;
        }

        [HttpGet]
        public HttpResponseMessage Status()
        {
            return new HttpResponseMessage
            {
                Content = new StringContent(
                    "{ Status: 'OK' }",
                    Encoding.UTF8,
                    "application/json"
                )
            };
        }


    }
}
