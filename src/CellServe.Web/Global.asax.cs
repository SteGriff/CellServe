using System.Web.Mvc;
using SimpleInjector;
using SimpleInjector.Integration.Web;
using SimpleInjector.Integration.Web.Mvc;

using System.Web.Optimization;
using System.Web.Routing;
using CellServe.ExcelHandler.Interfaces;
using CellServe.ExcelHandler;
using CellServe.ExcelHandler.Strategies;
using System.Reflection;
using LazyCache;
using CellServe.ExcelHandler.Adapters;

namespace CellServe.Web
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            // Set up SimpleInjector
            var container = new Container();
            container.Options.DefaultScopedLifestyle = new WebRequestLifestyle();

            container.Register<IWorkbookRepository, WorkbookRepository>(Lifestyle.Scoped);
            container.Register<ISheetFilterStrategy, SheetFilterStrategy>(Lifestyle.Scoped);
            container.Register<IRowModelingStrategy, RowModelingStrategy>(Lifestyle.Scoped);
            container.Register<IAppCache, AppCacheAdapter>(Lifestyle.Singleton);

            container.RegisterMvcControllers(Assembly.GetExecutingAssembly());
            container.Verify();

            DependencyResolver.SetResolver(new SimpleInjectorDependencyResolver(container));
        }
    }
}
