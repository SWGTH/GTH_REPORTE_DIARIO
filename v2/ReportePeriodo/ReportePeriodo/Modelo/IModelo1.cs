using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportePeriodo.Entidad;

namespace ReportePeriodo.Modelo
{
    public partial interface IModelo1
    {
        void DatosTeoricos(int ranId, int horaCorte, DateTime fechaIni, DateTime fechaFin, out List<DatosTeorico> listaProduccion,
            out List<DatosTeorico> listaSecas, out List<DatosTeorico> listaReto, out List<DatosTeorico> listaDestet1,
            out List<DatosTeorico> listaDestete2, out List<DatosTeorico> listaVaquillas, ref string mensaje);
        DatosProduccion DatosProduccion(DateTime inicio, DateTime fin, ref string mensaje);
        List<LecheBacteriologia> ListaLecheBacteriologia(DateTime inicio, DateTime fin, ref string mensaje);

    }
}
