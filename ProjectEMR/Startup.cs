using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ProjectEMR.Startup))]
namespace ProjectEMR
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
