using Microsoft.Owin;
using Owin;
using System.Web.Http;
using System.Web.Http.Cors;

[assembly: OwinStartupAttribute(typeof(DNA_CAPI_MIS.Startup))]
namespace DNA_CAPI_MIS
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            HttpConfiguration config = new HttpConfiguration();

            // Enable CORS globally
            var cors = new EnableCorsAttribute("*", "*", "*");
            config.EnableCors(cors);

            // Enable Web API
            WebApiConfig.Register(config);

            ConfigureAuth(app);
        }
    }
}
