using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entity
{
    public class Rancho
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
        public int EmpresaId { get; set; }
        public bool NoIdReal { get; set; }
        public bool VerLitros { get; set; }
        public bool Pesadores { get; set; }
    }
}
