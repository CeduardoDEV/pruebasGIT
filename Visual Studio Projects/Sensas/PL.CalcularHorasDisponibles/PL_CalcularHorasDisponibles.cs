using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PL.CalcularHorasDisponibles
{
    public class PL_CalcularHorasDisponibles : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService _service = factory.CreateOrganizationService(context.UserId);

            if (context.MessageName.ToUpper() == "UPDATE")
            {
                Entity PostImage;
                if (context.PostEntityImages.Contains("PostImage1") && context.PostEntityImages["PostImage1"] is Entity)
                {
                    PostImage = (Entity)context.PostEntityImages["PostImage1"];
                    
                }
            }
        }
    }
}
