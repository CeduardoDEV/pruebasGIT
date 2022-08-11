using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string serverUrl = "https://sensasdev.crm2.dynamics.com/";
            string clientId = "a73d011d-188a-4f98-9cce-f57a455e765b";
            string clientSecret = "wwN7Q~tbcplbzhQfHLA7T4y7KEomVRBtnM_6j";

            CrmServiceClient conn = new CrmServiceClient(new Uri(serverUrl), clientId, clientSecret, false, "");
            IOrganizationService _service = (IOrganizationService)conn.OrganizationWebProxyClient != null ?
            (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

            Program test = new Program();
            Guid detalleReporteId = new Guid("afb3e033-e941-ea11-a812-000d3ac04131");
            test.descontarHoras(_service, detalleReporteId);

        }
        public void descontarHoras(IOrganizationService service, Guid detalleId)
        {
            Entity detalleReporte = service.Retrieve("ce_detallereportedeservicio", detalleId, new ColumnSet("ce_reportedeservicios", "ce_duracion"));
            Entity reporteServicio = service.Retrieve("ce_reportedeservicio", ((EntityReference)detalleReporte.Attributes["ce_reportedeservicios"]).Id, new ColumnSet("ce_cliente"));
            Entity cliente = service.Retrieve("account", ((EntityReference)reporteServicio.Attributes["ce_cliente"]).Id, new ColumnSet("ce_horasprepagodisponibles"));
            Int32 duracionServicio = (Int32)detalleReporte.Attributes["ce_duracion"];
            decimal duracion = duracionServicio;
            decimal horasDisponibles = (decimal)cliente.Attributes["ce_horasprepagodisponibles"];
            decimal facturar = horasDisponibles - (duracion / 60);
            
            //Si el valor es menor a 0 afectar el valor del detalle
            if (facturar < 0)
            {
                facturar = facturar * 60;
                int facturarDetalle = (int)Math.Abs(facturar);
                cliente.Attributes["ce_horasprepagodisponibles"] = 0.0;
                detalleReporte.Attributes["ce_duracion"] = facturarDetalle;
                service.Update(detalleReporte);
                service.Update(cliente);
            }
            else 
            {
                cliente.Attributes["ce_horasprepagodisponibles"] = facturar;
                detalleReporte.Attributes["ce_duracion"] = 0;
                service.Update(detalleReporte);
                service.Update(cliente);
            }
            
        }
    }
}
