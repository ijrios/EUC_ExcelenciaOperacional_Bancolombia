using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Qualitas.Models
{
    public class ReprocesoPerdida
    {
        public int Id { get; set; }
        public string FechaOp { get; set; }
        public string FechaReg { get; set; }
        public string Consecutivo { get; set; }
        public string Nit { get; set; }
        public string Moneda { get; set; }
        public decimal Valor { get; set; }
        public string Cliente { get; set; }
        public string ProdEvento { get; set; }
        public string Responsable { get; set; }
        public string UsuarioReproceso { get; set; }
        public string Perdida { get; set; }
        public string Impacto { get; set; }
        public string Causa { get; set; }
        public string Descripcion { get; set; }
        public string Area { get; set; }
        public string TipoError { get; set; }
        public string Mes { get; set; }
    }
}
