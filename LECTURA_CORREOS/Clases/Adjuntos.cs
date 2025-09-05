using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LECTURA_CORREOS.Clases
{
    public class Adjuntos
    {
        public string nodesTpoDocRef { get; set; }
        public string nodesTipoDTE { get; set; }
        public string nodesFolio { get; set; }
        public string nodesFchEmis { get; set; }
        public string nodesRUTEmisor { get; set; }
        public string nodesRUTRecep { get; set; }
        public decimal nodesMntTotal { get; set; }
        public string nodesFolioRef { get; set; }
        public string nodesRazonRef { get; set; }
        public decimal nodesIVA { get; set; }
        public decimal nodesMntNeto { get; set; }
        public string TSTED { get; set; }
        public string rutaXML { get; set; }
        public string RznSocRecep { get; set; }
        public string nodeMntExe { get; set; }
        public string nodeRutReceptor { get; set; }
        public List<Referencias> listReferencias { get; set; }

    }
}
