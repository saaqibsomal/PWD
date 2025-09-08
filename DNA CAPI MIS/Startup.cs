using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(DNA_CAPI_MIS.Startup))]
namespace DNA_CAPI_MIS
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            app.UseCors(Microsoft.Owin.Cors.CorsOptions.AllowAll);
            ConfigureAuth(app);
        }
    }
}
