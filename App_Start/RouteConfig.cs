using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace Qualitas
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );



            routes.MapRoute(
                name: "Registro",
                url: "Home/Registro",
                defaults: new { controller = "Home", action = "Registro" }
            );

            routes.MapRoute(
                name: "Consultas",
                url: "Home/Consultas",
                defaults: new { controller = "Home", action = "Consultas" }
            );

            routes.MapRoute(
               name: "Administrador",
               url: "Home/Administrador",
               defaults: new { controller = "Home", action = "Administrador" }
           );

            routes.MapRoute(
               name: "Informes",
               url: "Home/Informes",
               defaults: new { controller = "Home", action = "Informes" }
           );

            routes.MapRoute(
               name: "ReporteTotal",
               url: "Home/ReporteTotal",
               defaults: new { controller = "Home", action = "ReporteTotal" }
           );

            routes.MapRoute(
               name: "Asesores",
               url: "Home/Asesores",
               defaults: new { controller = "Home", action = "Asesores" }
           );

            routes.MapRoute(
           name: "Perdidas",
           url: "Home/Perdidas",
           defaults: new { controller = "Home", action = "Perdidas" }
       );


            routes.MapRoute(
                name: "ObtenerReprocesos",
                url: "Home/ObtenerReprocesos",
               defaults: new { controller = "Home", action = "ObtenerReprocesos" }
            );
            routes.MapRoute(
               name: "ObtenerError",
               url: "Home/ObtenerError",
              defaults: new { controller = "Home", action = "ObtenerError" }
           );
            routes.MapRoute(
                name: "ErroresUsuarios",
                url: "Home/ErroresUsuarios",
               defaults: new { controller = "Home", action = "ErroresUsuarios" }
            );


            routes.MapRoute(
                name: "ObtenerReprocesosPerdidas",
                url: "Home/ObtenerReprocesosPerdidas",
               defaults: new { controller = "Home", action = "ObtenerReprocesosPerdidas" }
            );

            routes.MapRoute(
              name: "ObtenerErrorTipo",
              url: "Home/ObtenerErrorTipo",
             defaults: new { controller = "Home", action = "ObtenerErrorTipo" }
          );

            routes.MapRoute(
        name: "ObtenerErrorTipoImp",
        url: "Home/ObtenerErrorTipoImp",
       defaults: new { controller = "Home", action = "ObtenerErrorTipoImp" }
    );

            routes.MapRoute(
              name: "ObtenerErrorArea",
              url: "Home/ObtenerErrorArea",
             defaults: new { controller = "Home", action = "ObtenerErrorArea" }
          );


            routes.MapRoute(
            name: "ObtenerReprocesosNoPerdidas",
            url: "Home/ObtenerReprocesosNoPerdidas",
           defaults: new { controller = "Home", action = "ObtenerReprocesosNoPerdidas" }
        );

            routes.MapRoute(
               name: "ObtenerUsuario",
               url: "Home/ObtenerUsuario",
              defaults: new { controller = "Home", action = "ObtenerUsuario" }
           );


            routes.MapRoute(
               name: "GetCausasCumplimiento",
               url: "Home/GetCausasCumplimiento",
              defaults: new { controller = "Home", action = "GetCausasCumplimiento" }
           );

            routes.MapRoute(
             name: "GetCausasAprobacion",
             url: "Home/GetCausasAprobacion",
            defaults: new { controller = "Home", action = "GetCausasAprobacion" }
         );

            routes.MapRoute(
             name: "GetCausasOrientacion",
             url: "Home/GetCausasOrientacion",
            defaults: new { controller = "Home", action = "GetCausasOrientacion" }
         );

            routes.MapRoute(
             name: "GetCausasVerificacion",
             url: "Home/GetCausasVerificacion",
            defaults: new { controller = "Home", action = "GetCausasVerificacion" }
         );


            routes.MapRoute(
             name: "GetCausasAsignacion",
             url: "Home/GetCausasAsignacion",
            defaults: new { controller = "Home", action = "GetCausasAsignacion" }
         );


            routes.MapRoute(
               name: "AreasInfo",
               url: "Home/AreasInfo",
              defaults: new { controller = "Home", action = "AreasInfo" }
           );

            routes.MapRoute(
           name: "ImpactoInfo",
           url: "Home/ImpactoInfo",
          defaults: new { controller = "Home", action = "ImpactoInfo" }
       );


            routes.MapRoute(
               name: "Causas",
               url: "Home/Causas",
              defaults: new { controller = "Home", action = "Causas" }
           );


        }
    }
}
