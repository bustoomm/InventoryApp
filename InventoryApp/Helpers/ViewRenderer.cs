using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.AspNetCore.Mvc;

namespace InventoryApp.Helpers
{
    public static class ViewRenderer
    {
        public static string RenderViewToString<TModel>(this Controller controller, string viewName, TModel model)
        {
            var serviceProvider = controller.HttpContext.RequestServices;
            var engine = serviceProvider.GetService(typeof(ICompositeViewEngine)) as ICompositeViewEngine;
            var tempDataProvider = serviceProvider.GetService(typeof(ITempDataProvider)) as ITempDataProvider;

            var actionContext = new ActionContext(controller.HttpContext, controller.RouteData, controller.ControllerContext.ActionDescriptor);

            using var sw = new StringWriter();
            var viewResult = engine.FindView(actionContext, viewName, false);
            var viewContext = new ViewContext(actionContext, viewResult.View, new ViewDataDictionary<TModel>(
                new EmptyModelMetadataProvider(), new ModelStateDictionary())
            { Model = model }, new TempDataDictionary(controller.HttpContext, tempDataProvider), sw, new HtmlHelperOptions());

            viewResult.View.RenderAsync(viewContext).Wait();
            return sw.ToString();
        }
    }
}
