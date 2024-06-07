using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Qualitas.Models
{
    public class ErrorTotal
    {
        public string Usuario { get; set; }
        public string Errores { get; set; }
        public string Porcentaje { get; set; }
        public string Correo { get; set; }
        public string Mes { get; set; }
    }
}
