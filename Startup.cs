using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ReadExcelFile.Startup))]
namespace ReadExcelFile
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
