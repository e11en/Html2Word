using System.Web.Http;

namespace Html2Word
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            config.EnableCors();

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "Index",
                routeTemplate: "",
                defaults: new { controller = "Home", action = "Index" }
            );

            config.Routes.MapHttpRoute(
                name: "Status",
                routeTemplate: "api",
                defaults: new { controller = "Home", action = "Status" }
            );

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
