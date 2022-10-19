using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportePeriodo.Entidad;
using ReportePeriodo.Modelo;
using LibreriaAlimentacion;


namespace ReportePeriodo.Controlador
{
    public class Controlador1
    {
        Modelo1 modelo;

        public Controlador1()
        {
            modelo = new Modelo1();
        }

        public void DatosTeoricos(int ranId, int horaCorte, DateTime fechaIni, DateTime fechaFin, out List<DatosTeorico> listaProduccion,
            out List<DatosTeorico> listaSecas, out List<DatosTeorico> listaReto, out List<DatosTeorico> listaDestet1,
            out List<DatosTeorico> listaDestete2, out List<DatosTeorico> listaVaquillas, ref string mensaje)
        {
            listaProduccion = new List<DatosTeorico>();
            listaSecas = new List<DatosTeorico>();
            listaReto = new List<DatosTeorico>();
            listaDestet1 = new List<DatosTeorico>();
            listaDestete2 = new List<DatosTeorico>();
            listaVaquillas = new List<DatosTeorico>();
            mensaje = string.Empty;

            modelo.DatosTeoricos(ranId, horaCorte, fechaIni, fechaFin, out listaProduccion, out listaSecas, out listaReto, out listaDestet1, out listaDestete2, out listaVaquillas, ref mensaje);
        }

        public DatosProduccion DatosProduccion(DateTime inicio, DateTime fin, ref string mensaje)
        {
            return modelo.DatosProduccion(inicio, fin, ref mensaje);
        }

        public List<LecheBacteriologia> ListaLecheBacteriologia(DateTime inicio, DateTime fin, ref string mensaje)
        {
            return modelo.ListaLecheBacteriologia(inicio, fin, ref mensaje);
        }

        public bool NoIdReal(ref string mensaje)
        {
            return modelo.NoIdReal(ref mensaje);
        }

        public void CierreMesCorrecto(int ranId, int horaCorte,DateTime fechaInicio, DateTime fechaFin, out bool validacionMedicina, out bool validacionAlimento)
        {                        
            bool validacionCCO = modelo.CentrosDeCostoValidos(ranId, fechaInicio, fechaFin);
            string mensaje = string.Empty;
            GTH.ValidacionesCierreMes(ranId, horaCorte, fechaInicio, fechaFin, out validacionAlimento, out validacionMedicina, ref mensaje);
            bool auxiliarMedicina = validacionCCO && validacionMedicina;
            validacionMedicina = auxiliarMedicina;

        }


    }
}
