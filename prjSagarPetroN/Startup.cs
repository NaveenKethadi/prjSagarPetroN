using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(prjSagarPetroN.Startup))]
namespace prjSagarPetroN
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
