using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PL.ActualizarHorasDisponibles
{
    public class DatosIntegracion
    {
        public decimal horasDesc { get; set; }
        public string actividades { get; set; }
    }
    public class PL_ActualizarHorasDisponibles : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService _service = factory.CreateOrganizationService(context.UserId);

            if(context.MessageName.ToUpper() == "UPDATE") 
            {
                Entity PostImage;
                if(context.PostEntityImages.Contains("PostImage1") && context.PostEntityImages["PostImage1"] is Entity)
                {
                    PostImage = (Entity)context.PostEntityImages["PostImage1"];
                    Guid clienteId = ((EntityReference)PostImage.Attributes["ce_cliente"]).Id;

                    if (((OptionSetValue)PostImage.Attributes["ce_estado"]).Value == 100000001)
                    {
                        EntityCollection detalleReportes = obtenerDetalles(_service, PostImage.Id);
                        DatosIntegracion horasDescontar = descontarHoras(_service, detalleReportes, clienteId, tracingService);
                        Entity actualizarReporte = new Entity("ce_reportedeservicio", PostImage.Id);
                        actualizarReporte["sens_horasdescontadas"] = horasDescontar.horasDesc;
                        actualizarReporte["sens_actividadesrelacionadas"] = horasDescontar.actividades;
                        _service.Update(actualizarReporte);

                        EntityCollection detalleReportes2 = obtenerDetalles(_service, PostImage.Id);
                        datosERP(_service, detalleReportes2, actualizarReporte);
                    }
                    if (((OptionSetValue)PostImage.Attributes["ce_estado"]).Value == 100000003)
                    {
                        decimal horasRetornar = (decimal)PostImage.Attributes["sens_horasdescontadas"];
                        actualizarCuenta(_service,clienteId,horasRetornar);
                    }
                }
            }
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
        public DatosIntegracion descontarHoras(IOrganizationService service, EntityCollection detalleDeReporte, Guid cuentaId, ITracingService trace)
        {
            trace.Trace("Ingresa a la funcion descontarHoras");
            DatosIntegracion datosInte = new DatosIntegracion();
            decimal descontar = 0;
            string activi = "";

            foreach (Entity detalle in detalleDeReporte.Entities)
            {
                Entity cuenta = service.Retrieve("account", cuentaId, new ColumnSet("ce_horasprepagodisponibles", "ce_tipodeplizaprepagada"));
                trace.Trace("Ingresa al foreach");
                Int32 duracionServicio = (Int32)detalle.Attributes["ce_duracion"];
                trace.Trace("La duracion del detalle es: " + duracionServicio);
                string actividad = detalle.FormattedValues["ce_tipodeactividades"];
                decimal duracion = duracionServicio;
                trace.Trace(cuenta.Attributes["ce_horasprepagodisponibles"].ToString());
                decimal horasDisponibles = (decimal)cuenta.Attributes["ce_horasprepagodisponibles"];
                trace.Trace("Las horas disponibles son: " + horasDisponibles);
                decimal facturar = horasDisponibles - (duracion / 60);
                trace.Trace("El valor de la resta es: " + facturar);
                activi = actividad;

                if (((OptionSetValue)cuenta.Attributes["ce_tipodeplizaprepagada"]).Value == 134300000)
                {
                    if (facturar < 0)
                    { 
                        facturar = (int)Math.Round(facturar * 60);
                        int facturarDetalle = (int)Math.Abs(facturar);
                        trace.Trace("La conversion a minutos es: " + facturarDetalle);
                        cuenta.Attributes["ce_horasprepagodisponibles"] = 0.0;
                        detalle.Attributes["sens_tiempofacturar"] = facturarDetalle;
                        descontar += horasDisponibles;
                        service.Update(detalle);
                        trace.Trace("Se actualiza el detalle");
                        service.Update(cuenta);
                        trace.Trace("Se actualiza la cuenta");
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
                    detalle["sens_tiempofacturar"] = duracionServicio;
                    service.Update(detalle);
                }
            }
            trace.Trace("La cantidad total de horas a descontar es " + descontar);
            datosInte.horasDesc = descontar;
            datosInte.actividades = activi;
            //datosInte.actividades = String.Join(", ", actividadesLista);
            trace.Trace("termina funcion descontarHoras");
            return datosInte;
        }
        public void actualizarCuenta(IOrganizationService service, Guid cuentaId, decimal horas)
        {
            Entity cuenta = service.Retrieve("account", cuentaId, new ColumnSet("ce_horasprepagodisponibles", "ce_tipodeplizaprepagada"));
            cuenta.Attributes["ce_horasprepagodisponibles"] = horas + (decimal)cuenta.Attributes["ce_horasprepagodisponibles"];
            service.Update(cuenta);
        }
        public void datosERP(IOrganizationService service, EntityCollection detalleDeReporte, Entity reporteA)
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
