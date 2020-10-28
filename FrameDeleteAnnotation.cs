using Microsoft.Xrm.Sdk;
using System;

namespace Microsoft.Crm.Sdk.Samples
{
    public class FrameDeleteAnnotation : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            service.Delete(context.PrimaryEntityName, context.PrimaryEntityId);
        }
    }
}
