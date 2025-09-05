using LECTURA_CORREOS.Clases;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Xml;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Data.Common;

namespace LECTURA_CORREOS
{
    public class Program
    {
        static SqlConnection conexionBaseDatos;
        static async Task Main(string[] args)
        {
            leerDatos(@"C:\Users\Usuario\Pictures\76249099-4_20250728_163605.xml", 0);
            // Lee el contenido del archivo JSON
            string rutaJson = ConfigurationManager.AppSettings["rutaConfiguracion"];
            string json = File.ReadAllText(rutaJson);
            JObject data = JObject.Parse(json);
            string rutaDescargaXmls = data["rutaDescargaXmls"].ToString();

            // Utilizar DbConnectionStringBuilder para analizar la cadena de conexión
            var builder = new DbConnectionStringBuilder();
            builder.ConnectionString = data["conexionBaseDatos"].ToString();

            // Extraer los valores específicos
            string servidor = builder.ContainsKey("Server") ? builder["Server"].ToString() : "";
            string baseDatos = builder.ContainsKey("Database") ? builder["Database"].ToString() : "";
            string usuario = builder.ContainsKey("UID") ? builder["UID"].ToString() : "";
            string password = builder.ContainsKey("PWD") ? builder["PWD"].ToString() : "";



            conexionBaseDatos = BaseDatos.abrirConexionSQL(servidor, baseDatos, usuario, password);
            //conexionBaseDatos = BaseDatos.abrirConexionSQL("GFSERVER05", "RPA_FACTURAS_SIED", "sa", "Softland363");

            if (conexionBaseDatos != null)
            {
                while (true)
                {
                    Console.WriteLine("\n --------------- SENSANDO BANDEJA DE CORREO ELECTRÓNICO ----------------");
                    await Task.Delay(5000);
                    string fechaInicioLecturaCorreo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    List<Correo> correos = lecturaCorreos(rutaDescargaXmls);
                    string fechaFinLecturaCorreo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    if (correos.Count > 0)
                    {
                        ingresarCorreosBaseDatos(correos, conexionBaseDatos, fechaInicioLecturaCorreo, fechaFinLecturaCorreo);
                    }
                }
            }
        }


        static List<Correo> lecturaCorreos(string rutaDescargaXml)
        {
            List<Correo> listCorreo = new List<Correo>();

            // Crear una instancia de la aplicación de Outlook
            Outlook.Application outlookApp = new Outlook.Application();

            // Obtener la carpeta de la bandeja de entrada
            Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            // Crear un filtro para obtener sólo los correos no leídos
            string filter = "[UnRead]=true";
            Outlook.Items unreadItems = inboxFolder.Items.Restrict(filter);


            int index = 0;
            int cantidadCorreos = unreadItems.Count;

            // Recorrer los elementos no leídos
            foreach (object item in unreadItems)
            {
                index++;
                Console.WriteLine("");
                Console.WriteLine("-------------------- CORREO " + index.ToString() + " DE " + cantidadCorreos + " --------------------");
                
                if (item is Outlook.MailItem mail)
                {
                    try
                    {
                        Console.WriteLine("IDENTIFICADOR :    " + mail.EntryID);
                        Console.WriteLine("TITULO :           " + mail.Subject);
                        Console.WriteLine("DE :               " + mail.SenderName);
                        Console.WriteLine("RECIVIDO :         " + mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"));
                        Console.WriteLine("ARCHIVOS ADJUNTO : " + mail.Attachments.Count.ToString());
                        // Recorrer los archivos adjuntos
                        if (mail.Attachments.Count > 0)
                        {
                            bool tieneXML = HasXmlAttachment(mail);
                            if (tieneXML)
                            {
                                List<Adjuntos> listXml = new List<Adjuntos>();
                                foreach (Outlook.Attachment attachment in mail.Attachments)
                                {
                                    if (attachment.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string horaActual = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                                        // Guardar el archivo XML en una ubicación temporal
                                        string tempFilePath = Path.GetTempFileName();
                                        attachment.SaveAsFile(tempFilePath);

                                        int cantidadDocumentos = verificarCantidadDocumentosXML(tempFilePath);
                                        for (int i = 0; i < cantidadDocumentos; i++)
                                        {
                                            bool existe = CheckIfTagExists(tempFilePath, i);

                                            if (existe)
                                            {
                                                // Leer el contenido del archivo XML
                                                Adjuntos datosXML = new Adjuntos();
                                                datosXML = leerDatos(tempFilePath, i);
                                                // Opcional: Guardar el archivo adjunto
                                                Thread.Sleep(4000);

                                                string savePath = obtenerRutaDescargaXML(rutaDescargaXml) + @"\" + datosXML.nodesRUTEmisor + "_" + horaActual + ".xml";
                                                if (!File.Exists(savePath))
                                                {
                                                    attachment.SaveAsFile(savePath);
                                                }
                                                datosXML.rutaXML = savePath;
                                                listXml.Add(datosXML);
                                            }
                                        }
                                        // Eliminar el archivo temporal si ya no se necesita
                                        File.Delete(tempFilePath);
                                    }
                                }

                                if (listXml.Count > 0)
                                {
                                    listCorreo.Add(new Correo()
                                    {
                                        idCorreo = mail.EntryID,
                                        correo = mail,
                                        de = mail.SenderName,
                                        titulo = mail.Subject,
                                        fechaCorreo = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"),
                                        adjuntos = listXml
                                    });
                                }
                                else
                                {
                                    // NO EXISTE ARCHIVOS .XML
                                    Console.WriteLine("OBSERVACIÓN : LOS .XML NO TIENEN EL FORMATO CORRECTO - CORREO MARCADO COMO LEIDO");
                                    mail.UnRead = false;
                                    mail.Save();
                                }
                            }
                            else
                            {
                                // NO EXISTE ARCHIVOS .XML
                                Console.WriteLine("OBSERVACIÓN : CORREO SIN ARCHIVOS .XML - CORREO MARCADO COMO LEIDO");
                                mail.UnRead = false;
                                mail.Save();
                            }
                        }
                        else
                        {
                            // NO EXISTE ADJUNTO EN EL CORREO
                            Console.WriteLine("OBSERVACIÓN : CORREO SIN ADJUNTOS - CORREO MARCADO COMO LEIDO");
                            mail.UnRead = false;
                            mail.Save();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine("OBSERVACIÓN : ERROR AL LEER CORREO - "+ex.Message);
                    }
                }
                else if (item is Outlook.AppointmentItem appointment)
                {
                    Console.WriteLine("OBSERVACIÓN : Evento de Calendario - CORREO MARCADO COMO LEIDO");
                    appointment.UnRead = false;
                    appointment.Save();
                }
                else if (item is Outlook.TaskItem task)
                {
                    Console.WriteLine("OBSERVACIÓN : Evento de Tarea - CORREO MARCADO COMO LEIDO");
                    task.UnRead = false;
                    task.Save();
                }
                else if (item is Outlook.ReportItem report)
                {
                    Console.WriteLine("OBSERVACIÓN : Notificación de No Entregado (NDR) - CORREO MARCADO COMO LEIDO");
                    report.UnRead = false;
                    report.Save();
                }
                else if (item is Outlook.MeetingItem meeting)
                {
                    Console.WriteLine("OBSERVACIÓN : Notificación de reuniones - CORREO MARCADO COMO LEIDO");
                    meeting.UnRead = false;
                    meeting.Save();
                }
            }
            // Liberar los objetos COM
            Marshal.ReleaseComObject(inboxFolder);
            Marshal.ReleaseComObject(outlookNamespace);
            Marshal.ReleaseComObject(outlookApp);
            return listCorreo;
        }

        // Función para verificar si un correo tiene al menos un archivo .xml adjunto
        static bool HasXmlAttachment(Outlook.MailItem mail)
        {
            foreach (Outlook.Attachment attachment in mail.Attachments)
            {
                if (attachment.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
        // Función para leer y procesar el archivo XML
        static Adjuntos leerDatos(string filePath,int numeroDocumento)
        {
            Adjuntos datos = new Adjuntos();
            List<Referencias> referencias = new List<Referencias>();
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(filePath);
                XmlNodeList dteNodes = doc.GetElementsByTagName("DTE");
                XmlNodeList dteNodesCaratula = doc.GetElementsByTagName("Caratula");
                XmlElement dteElement = (XmlElement)dteNodes[numeroDocumento];
                XmlElement dteElementCaratula = (XmlElement)dteNodesCaratula[0];


                XmlNodeList nodesTpoDocRef = dteElement.GetElementsByTagName("TpoDocRef");
                XmlNodeList nodesTipoDTE = dteElement.GetElementsByTagName("TipoDTE");
                XmlNodeList nodesFolio = dteElement.GetElementsByTagName("Folio");
                XmlNodeList nodesFchEmis = dteElement.GetElementsByTagName("FchEmis");
                XmlNodeList nodesRUTEmisor = dteElement.GetElementsByTagName("RUTEmisor");
                XmlNodeList nodesRUTRecep = dteElement.GetElementsByTagName("RUTRecep");
                XmlNodeList nodesMntTotal = dteElement.GetElementsByTagName("MntTotal");
                XmlNodeList nodesFolioRef = dteElement.GetElementsByTagName("FolioRef");
                XmlNodeList nodesRazonRef = dteElement.GetElementsByTagName("RazonRef");
                XmlNodeList nodesIVA = dteElement.GetElementsByTagName("IVA");
                XmlNodeList nodesMntNeto = dteElement.GetElementsByTagName("MntNeto");
                XmlNodeList nodeTSTED = dteElement.GetElementsByTagName("TSTED");
                XmlNodeList nodesRznSocRecep = dteElement.GetElementsByTagName("RznSocRecep");
                XmlNodeList nodesMntExe = dteElement.GetElementsByTagName("MntExe");
                XmlNodeList nodesReferencias = dteElement.SelectNodes("//*[local-name()='Referencia']");



                try
                {
                    datos.nodeMntExe = nodesMntExe[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodeMntExe = "";
                }

                try
                {
                    datos.RznSocRecep = nodesRznSocRecep[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.RznSocRecep = "";
                }


                try
                {
                    datos.nodesTpoDocRef = nodesTpoDocRef[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesTpoDocRef = "";
                }

                try
                {
                    datos.nodesTipoDTE = nodesTipoDTE[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesTipoDTE = "";
                }

                try
                {
                    datos.nodesFolio = nodesFolio[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesFolio = "";
                }

                try
                {
                    datos.nodesFchEmis = nodesFchEmis[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesFchEmis = "";
                }

                try
                {
                    datos.nodesRUTEmisor = nodesRUTEmisor[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesRUTEmisor = "";
                }

                try
                {
                    datos.nodesRUTRecep = nodesRUTRecep[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesRUTRecep = "";
                }

                try
                {
                    datos.nodesMntTotal = Convert.ToDecimal(nodesMntTotal[0].InnerText.Trim());
                }
                catch (System.Exception ex)
                {
                    datos.nodesMntTotal = 0;
                }

                try
                {
                    datos.nodesFolioRef = nodesFolioRef[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesFolioRef = "";
                }

                try
                {
                    datos.nodesRazonRef = nodesRazonRef[0].InnerText.Trim();
                }
                catch (System.Exception ex)
                {
                    datos.nodesRazonRef = "";
                }

                try
                {
                    datos.nodesIVA = Convert.ToDecimal(nodesIVA[0].InnerText.Trim());
                }
                catch (System.Exception ex)
                {
                    datos.nodesIVA = 0;
                }

                try
                {
                    datos.nodesMntNeto = Convert.ToDecimal(nodesMntNeto[0].InnerText.Trim());
                }
                catch (System.Exception ex)
                {
                    datos.nodesMntNeto = 0;
                }

                try
                {
                    datos.TSTED = nodeTSTED[0].InnerText.Trim();
                }
                catch (System.Exception)
                {
                    datos.TSTED = "";
                }

                try
                {
                    XmlNodeList nodesRutReceptor = dteElementCaratula.GetElementsByTagName("RutReceptor");
                    datos.nodeRutReceptor = nodesRutReceptor[0].InnerText.Trim();
                }
                catch (System.Exception)
                {
                    datos.nodeRutReceptor = "";
                }

                foreach (XmlNode item in nodesReferencias)
                {
                    string TpoDocRef = "";
                    string FolioRef = "";
                    string FchRef = "";
                    try
                    {
                        XmlNode nodoTpoDocRef = item.SelectSingleNode("*[local-name()='TpoDocRef']");
                        if (nodoTpoDocRef != null) TpoDocRef = nodoTpoDocRef.InnerText;
                    }
                    catch (System.Exception ex)
                    {
                    }
                    try
                    {
                        XmlNode nodoFolioRef = item.SelectSingleNode("*[local-name()='FolioRef']");
                        if (nodoFolioRef != null) FolioRef = nodoFolioRef.InnerText;
                    }
                    catch (System.Exception ex)
                    {
                    }
                    try
                    {
                        XmlNode nodoFchRef = item.SelectSingleNode("*[local-name()='FchRef']");
                        if (nodoFchRef != null) FchRef = nodoFchRef.InnerText;
                    }
                    catch (System.Exception ex)
                    {
                    }
                    referencias.Add(new Referencias
                    {
                        TpoDocRef = TpoDocRef,
                        FolioRef = FolioRef,
                        FchRef = FchRef,
                    });
                }

            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Error al leer el archivo XML: " + ex.Message);
            }
            datos.listReferencias = referencias;
            return datos;
        }
        static bool CheckIfTagExists(string filePath, int numeroDocumento)
        {
            bool exists = false;
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(filePath);
                XmlNodeList dteNodes = doc.GetElementsByTagName("DTE");
                XmlElement dteElement = (XmlElement)dteNodes[numeroDocumento];


                XmlNodeList nodesTpoDocRef = dteElement.GetElementsByTagName("TpoDocRef");
                XmlNodeList nodesTipoDTE = dteElement.GetElementsByTagName("TipoDTE");
                XmlNodeList nodesFolio = dteElement.GetElementsByTagName("Folio");
                XmlNodeList nodesFchEmis = dteElement.GetElementsByTagName("FchEmis");
                XmlNodeList nodesRUTEmisor = dteElement.GetElementsByTagName("RUTEmisor");
                XmlNodeList nodesRUTRecep = dteElement.GetElementsByTagName("RUTRecep");
                XmlNodeList nodesMntTotal = dteElement.GetElementsByTagName("MntTotal");
                XmlNodeList nodesFolioRef = dteElement.GetElementsByTagName("FolioRef");
                XmlNodeList nodesRazonRef = dteElement.GetElementsByTagName("RazonRef");
                XmlNodeList nodesIVA = dteElement.GetElementsByTagName("IVA");
                XmlNodeList nodesMntNeto = dteElement.GetElementsByTagName("MntNeto");
                XmlNodeList nodesTSTED = dteElement.GetElementsByTagName("TSTED");
                XmlNodeList nodesRznSocRecep = dteElement.GetElementsByTagName("RznSocRecep");
                XmlNodeList nodesMntExe = dteElement.GetElementsByTagName("MntExe");

                if (nodesTipoDTE.Count > 0 & nodesTSTED.Count >0 & nodesFolio.Count > 0 & nodesFchEmis.Count > 0 & nodesRUTEmisor.Count >0 &
                    nodesRUTRecep.Count >0
                   )
                {
                    string tipoDTE = nodesTipoDTE[0].InnerText.Trim();
                    if (tipoDTE == "34")
                    {
                        if (nodesMntTotal.Count > 0 & nodesMntExe.Count > 0)
                        {
                            exists = true;
                        }
                    }
                    else
                    {
                        if (nodesMntNeto.Count > 0 & nodesIVA.Count > 0 & nodesMntTotal.Count > 0 & nodesRznSocRecep.Count > 0)
                        {
                            if (nodesTpoDocRef.Count > 0)
                            {
                                string tpoDocRef = nodesTpoDocRef[0].InnerText.Trim();
                                if (tpoDocRef == "SEN" & (tipoDTE == "33" | tipoDTE == "61"))
                                {
                                    if (nodesFolioRef.Count > 0 & nodesRazonRef.Count > 0)
                                    {
                                        exists = true;
                                    }
                                    else
                                    {
                                        exists = false;
                                    }
                                }
                                else
                                {
                                    exists = true;
                                }
                            }
                            else
                            {
                                exists = true;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Error al leer el archivo XML: " + ex.Message);
            }
            return exists;
        }

        static string obtenerRutaDescargaXML(string ruta)
        {
            //string ruta = @"C:\GAIA\REGISTRO_FACTURAS";
            //ruta += @"\DESCARGAS_XML";
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            int anoActual = DateTime.Now.Year;
            string mes = DateTime.Now.ToString("MMMM").ToUpper();
            ruta += @"\"+anoActual.ToString();
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            ruta += @"\" + mes.ToString();
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            return ruta;
        }

        static void ingresarCorreosBaseDatos(List<Correo> correos,SqlConnection connection,string fechaInicioLecturaCorreo,string fechaFinLecturaCorreo)
        {
            int cantidadCorreos = correos.Count;
            int contadorCorreos = 0;
            foreach (Correo item in correos)
            {
                contadorCorreos++;
                Console.WriteLine("");
                Console.WriteLine("------------INGRESAR "+ contadorCorreos.ToString() + " DE "+ cantidadCorreos.ToString() + " A LA BASE DE DATOS" );
                Console.WriteLine("TITULO :           " + item.titulo);
                Console.WriteLine("DE :               " + item.de);
                Console.WriteLine("RECIVIDO :         " + item.fechaCorreo);
                Console.WriteLine("ARCHIVOS ADJUNTO : " + item.adjuntos.Count.ToString());

                int idCorreo = BaseDatos.ingresarCorreo(connection, item.idCorreo, item.titulo, item.de, item.fechaCorreo, fechaInicioLecturaCorreo, fechaFinLecturaCorreo);
                if (idCorreo != 0)
                {
                    try
                    {
                        bool marcarCorreoLeido = true;
                        int cantidadArchivos = 0;
                        foreach (Adjuntos archivosXML in item.adjuntos)
                        {
                            cantidadArchivos++;
                            Console.WriteLine("INGRESANDO " + cantidadArchivos.ToString() +" DE "+item.adjuntos.Count.ToString() + " ARCHIVOS ");
                            if (archivosXML.nodeMntExe == "")
                            {
                                archivosXML.nodeMntExe = "0";
                            }
                            int idFactura = BaseDatos.ingresarAdjuntoCorreo(connection, idCorreo, archivosXML.rutaXML, archivosXML.nodesTpoDocRef, archivosXML.nodesTipoDTE, archivosXML.nodesFolio, archivosXML.nodesFchEmis, archivosXML.nodesRUTEmisor, archivosXML.nodesRUTRecep, archivosXML.nodesMntTotal, archivosXML.nodesIVA, archivosXML.nodesMntNeto, archivosXML.nodesFolioRef, archivosXML.nodesRazonRef, archivosXML.TSTED, archivosXML.RznSocRecep, archivosXML.nodeMntExe, archivosXML.nodeRutReceptor);
                            Console.WriteLine("idFactura = "+ idFactura.ToString());
                            if (idFactura == 0)
                            {
                                marcarCorreoLeido = false;
                            }
                            else
                            {
                                foreach (Referencias itemReferencia in archivosXML.listReferencias)
                                {
                                    BaseDatos.ingresarReferencia(connection, idFactura, itemReferencia.TpoDocRef, itemReferencia.FolioRef, itemReferencia.FchRef);
                                }
                            }
                        }
                        if (marcarCorreoLeido == true)
                        {
                            item.correo.UnRead = false;
                            item.correo.Save();
                        }
                    }
                    catch (System.Exception)
                    {
                        Console.WriteLine("Error al ingresar xml del correo "+ item.titulo);
                    }
                }
            }
        }

        static int verificarCantidadDocumentosXML(string filePath)
        {
            int cantidadDocumentos = 0;
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(filePath);

                XmlNodeList nodesDTE = doc.GetElementsByTagName("DTE");
                cantidadDocumentos = nodesDTE.Count;
            }
            catch (System.Exception ex)
            {

            }
            return cantidadDocumentos;
        }


    }
}
