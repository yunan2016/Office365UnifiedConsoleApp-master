using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(OneDriveWeb.Startup))]
namespace OneDriveWeb
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
