using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class Rancho
    {
        public int Ran_ID { get; set; }
        public string Ran_Nombre { get; set; }
        public string Erp { get; set; }
        public int TimeShiftTracker { get; set; }
        public int Emp_ID { get; set; }
        public bool No_ID_Real { get; set; }
        public bool Ver_Litros { get; set; }
        public bool Pesadores { get; set; }
    }
}
