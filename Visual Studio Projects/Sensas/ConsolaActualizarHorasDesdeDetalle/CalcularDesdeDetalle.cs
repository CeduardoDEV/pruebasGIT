using ConsolaActualizarHorasDesdeDetalle;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
            Guid ReporteId = new Guid("c1902eba-dafe-ec11-82e5-000d3a889c6f");
            Entity reporte = _service.Retrieve("ce_reportedeservicio", ReporteId, new ColumnSet("sens_horasdescontadas", "sens_actividadesrelacionadas"));
            Guid clienteId = new Guid("57a924db-e933-ea11-a813-000d3ac04fef");
            EntityCollection detalleReportes =  test.obtenerDetalles(_service, ReporteId);
            
            DatosIntegracion horasDescontar = test.descontarHoras(_service, detalleReportes, clienteId);
            reporte.Attributes["sens_horasdescontadas"] = horasDescontar.horasDesc;
            reporte.Attributes["sens_actividadesrelacionadas"] = horasDescontar.actividades;
            _service.Update(reporte);
            
            EntityCollection detalleReportes2 = test.obtenerDetalles(_service, ReporteId);
            test.horasERP(_service, detalleReportes2, reporte);

            decimal retornarHoras = (decimal)reporte.Attributes["sens_horasdescontadas"];
            Entity cuentaA = _service.Retrieve("account", clienteId, new ColumnSet("ce_horasprepagodisponibles"));
            cuentaA.Attributes["ce_horasprepagodisponibles"] = retornarHoras + (decimal)cuentaA.Attributes["ce_horasprepagodisponibles"];
            _service.Update(cuentaA);

        }
        public EntityCollection obtenerDetalles(IOrganizationService service, Guid reporte)
        {
            string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='ce_detallereportedeservicio'>
                                    <attribute name='ce_detallereportedeservicioid' />
                                    <attribute name='ce_facturable' />
                                    <attribute name='ce_duracion' />
                                    <attribute name='sens_tiempofacturar' />
                                    <attribute name='ce_tipodeactividades' />
                                    <filter type='and'>
                                      <condition attribute='ce_reportedeservicios' operator='eq' uitype='ce_reportedeservicio' value='{reporte}' />
                                      <condition attribute='ce_facturable' operator='eq' value='1' />
                                    </filter>
                                  </entity>
                                </fetch>";

            EntityCollection detalles = service.RetrieveMultiple(new FetchExpression(query));
            return detalles;
        }
        public DatosIntegracion descontarHoras(IOrganizationService service, EntityCollection detalleDeReporte, Guid cuentaId)
        {
            var datosInte = new DatosIntegracion();

            decimal descontar = 0;
            //StringBuilder actividadesDetalle = new StringBuilder();
            string[] activiVacia = { };
            List<string> actividadesLista = new List<string>(activiVacia.ToList());

            foreach(Entity detalle in detalleDeReporte.Entities)
            {
                Entity cuenta = service.Retrieve("account", cuentaId, new ColumnSet("ce_horasprepagodisponibles", "ce_tipodeplizaprepagada"));
                string actividad = detalle.FormattedValues["ce_tipodeactividades"];
                Int32 duracionServicio = (Int32)detalle.Attributes["ce_duracion"];
                decimal duracion = duracionServicio;
                decimal horasDisponibles = (decimal)cuenta.Attributes["ce_horasprepagodisponibles"];
                decimal facturar = horasDisponibles - (duracion / 60);

                actividadesLista.Add(actividad);

                //Se valida si la cuenta es para Descontar horas "Horas Prepagadas Servicio"
                if (((OptionSetValue)cuenta.Attributes["ce_tipodeplizaprepagada"]).Value == 134300000)
                {
                    if (facturar < 0)
                    {
                        //facturar = (int)Math.Round(facturar * 60);
                        int facturarDetalle = (int)Math.Abs(facturar * 60);
                        //int facturarDetalle = (int)Math.Abs(facturar);
                        cuenta.Attributes["ce_horasprepagodisponibles"] = 0.0;
                        detalle.Attributes["sens_tiempofacturar"] = facturarDetalle;
                        descontar += horasDisponibles;
                        service.Update(detalle);
                        service.Update(cuenta);
                    }
                    else
                    {
                        cuenta.Attributes["ce_horasprepagodisponibles"] = facturar;
                        detalle.Attributes["sens_tiempofacturar"] = 0;
                        descontar += duracion / 60;
                        service.Update(detalle);
                        service.Update(cuenta);
                    }
                }
                else
                {
                    descontar += duracion;
                    detalle["sens_tiempofacturar"] = duracionServicio;
                    service.Update(detalle);
                }
            }
            datosInte.horasDesc = descontar;
            datosInte.actividades = String.Join(", ", actividadesLista);
            return datosInte;
        }
        public void horasERP(IOrganizationService service, EntityCollection detalleDeReporte, Entity reporteA)
        {
            int duracion = 0;

            foreach (Entity detalle in detalleDeReporte.Entities)
            {
                int duracionServicio = (int)detalle.Attributes["sens_tiempofacturar"];
                duracion += duracionServicio; 
            }

            reporteA["sens_tiempofacturadoerp"] = duracion;
            service.Update(reporteA);

        }
    }
}
