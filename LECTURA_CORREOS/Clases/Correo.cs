using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace LECTURA_CORREOS.Clases
{
    public class Correo
    {
        public string idCorreo { get; set; }
        public string titulo { get; set; }
        public string de { get; set; }
        public string fechaCorreo { get; set; }
        public Outlook.MailItem correo { get; set; }
        public List<Adjuntos> adjuntos { get; set; }
    }
}
