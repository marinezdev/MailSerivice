using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Text;

namespace WebServiceCorreo
{
    public class Correos
    {
        public string CorreoMetPrivadoEnvia(string _host, int _puerto, string _usuario, string _contra, string _Email, string _from, string _Subject, string _body)
        {
            string respuesta = "";

            /*
            string host = "mail.asae.com.mx";
            int puerto = 25;
            string usuario = "asae_contactanos@asae.com.mx";
            string contra = "A$ae2018$$_";
            */

            string host = _host;
            int puerto = _puerto;
            string usuario = _usuario;
            string contra = _contra;
            string pCorreo = _Email;
            string from = _from;
            string Subject = _Subject;
            string body = _body;

            MailMessage correo = new MailMessage();
            correo.To.Add(new MailAddress(pCorreo));
            correo.From = new MailAddress(usuario, from);
            correo.Subject = Subject;
            correo.Body = body;
            
            correo.IsBodyHtml = true;
            correo.Priority = MailPriority.Normal;

            SmtpClient smtp = new SmtpClient(host, puerto);
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(usuario, contra);

            ServicePointManager.ServerCertificateValidationCallback =
            delegate (object s
                , System.Security.Cryptography.X509Certificates.X509Certificate certificate
                , System.Security.Cryptography.X509Certificates.X509Chain chai
                , SslPolicyErrors sslPolicyErrors)
            {
                return true;
            };

            try
            {
                smtp.Send(correo);
                correo.Dispose();
                respuesta = "ok";
            }
            catch (Exception ex)
            {
                respuesta = "Error: " + ex.Message;
            }

            return respuesta;

        }

        public string CorreoAsaeTicketsEnvia(string _host, int _puerto, string _usuario, string _contra, string _Email, string _EmailCopia, string _from, string _Subject, string _body)
        {
            string respuesta = "";

            string host = _host;
            int puerto = _puerto;
            string usuario = _usuario;
            string contra = _contra;
            string pCorreo = _Email;
            string pCorreoCopia = _EmailCopia;
            string from = _from;
            string Subject = _Subject;
            string body = _body;

            MailMessage correo = new MailMessage();
            correo.To.Add(new MailAddress(pCorreo));
            if (pCorreoCopia.ToString() != String.Empty)
            {
                correo.CC.Add(new MailAddress(pCorreoCopia));
            }
                
            correo.From = new MailAddress(usuario, from);
            correo.Subject = Subject;
            correo.Body = body;

            correo.IsBodyHtml = true;
            correo.Priority = MailPriority.Normal;

            SmtpClient smtp = new SmtpClient(host, puerto);
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(usuario, contra);

            ServicePointManager.ServerCertificateValidationCallback =
            delegate (object s
                , System.Security.Cryptography.X509Certificates.X509Certificate certificate
                , System.Security.Cryptography.X509Certificates.X509Chain chai
                , SslPolicyErrors sslPolicyErrors)
            {
                return true;
            };

            try
            {
                smtp.Send(correo);
                correo.Dispose();
                respuesta = "ok";
            }
            catch (Exception ex)
            {
                respuesta = "Error: " + ex.Message;
            }

            return respuesta;

        }

        public string CorreoAsaeCRM_AgendaEvento(string _host, int _puerto, string _usuario, string _contra, string _Email, string _EmailCopia, string _from, string _Subject, string _body, string archivoNombre, string EventoTitulo, DateTime EventoFechaI, DateTime EventoFechaF, string EventoOrganizaNombre, string EventoOrganizaMail, string EventoAsunto, string EventoDescripcion, string EventoUbicacion)
        {
            string rutaArchivoEvento = @"c:\temporaly\";
            string respuesta = "";


            string host = _host;
            int puerto = _puerto;
            string usuario = _usuario;
            string contra = _contra;
            string pCorreo = _Email;
            string pCorreoCopia = _EmailCopia;
            string from = _from;
            string Subject = _Subject;
            string body = _body;

            rutaArchivoEvento += archivoNombre;
            Calenario_AgendaAdd(archivoNombre, EventoTitulo, EventoFechaI, EventoFechaF, EventoOrganizaNombre, EventoOrganizaMail, EventoAsunto, EventoDescripcion, EventoUbicacion);
            if (!File.Exists(rutaArchivoEvento) )
            {
                respuesta = "Error: No se pudó crear el archivo para agendar evento.";
                return respuesta;
            }


            MailMessage correo = new MailMessage();
            correo.To.Add(new MailAddress(pCorreo));
            if (pCorreoCopia.ToString() != String.Empty)
            {
                correo.CC.Add(new MailAddress(pCorreoCopia));
            }

            correo.From = new MailAddress(usuario, from);
            correo.Subject = Subject;
            correo.Body = body;

            correo.IsBodyHtml = true;
            correo.Priority = MailPriority.Normal;

            correo.Attachments.Add(new Attachment(rutaArchivoEvento));

            SmtpClient smtp = new SmtpClient(host, puerto);
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(usuario, contra);

            ServicePointManager.ServerCertificateValidationCallback =
            delegate (object s
                , System.Security.Cryptography.X509Certificates.X509Certificate certificate
                , System.Security.Cryptography.X509Certificates.X509Chain chai
                , SslPolicyErrors sslPolicyErrors)
            {
                return true;
            };

            try
            {
                smtp.Send(correo);
                correo.Dispose();
                respuesta = "ok";
            }
            catch (Exception ex)
            {
                respuesta = "Error: " + ex.Message;
            }

            try
            {
                string rutaArchivo = @"C:\temporaly\" + archivoNombre;
                File.Delete(rutaArchivo);
            }
            catch
            { }

            return respuesta;

        }

        private bool Calenario_AgendaAdd(string archivoNombre, string EventoTitulo, DateTime EventoFechaI, DateTime EventoFechaF, string EventoOrganizaNombre, string EventoOrganizaMail, string EventoAsunto, string EventoDescripcion, string EventoUbicacion)
        {
            bool _resultado = false;
            string rutaArchivo = @"C:\temporaly\" + archivoNombre;

            try
            {
                if (System.IO.File.Exists(archivoNombre))
                {
                    System.IO.File.Delete(archivoNombre);
                }

                FileStream archivo = new FileStream(rutaArchivo, FileMode.Append);
                using (StreamWriter file = new StreamWriter(archivo))
                {
                    file.WriteLine("BEGIN:VCALENDAR");
                    file.WriteLine("VERSION:2.0");
                    file.WriteLine("PRODID: -//hacksw/handcal//NONSGML v1.0//EN");
                    file.WriteLine("BEGIN:VEVENT");
                    file.WriteLine("UID:" + EventoTitulo);
                    file.WriteLine("DTSTAMP:" + EventoFechaI.ToString("yyyyMMddTHHmmss"));
                    file.WriteLine("ORGANIZER;CN=" + EventoOrganizaNombre + " MAILTO:" + EventoOrganizaMail);
                    file.WriteLine("DTSTART:" + EventoFechaI.ToString("yyyyMMddTHHmmss"));
                    file.WriteLine("DTEND:" + EventoFechaF.ToString("yyyyMMddTHHmmss"));
                    file.WriteLine("SUMMARY:" + EventoAsunto);
                    file.WriteLine("DESCRIPTION:" + EventoDescripcion);
                    file.WriteLine("LOCATION:" + EventoUbicacion);
                    file.WriteLine("END:VEVENT");
                    file.WriteLine("END:VCALENDAR");
                    _resultado = true;
                }
            }
            catch
            {
                _resultado = false;
            }

            return _resultado;
        }
    }
}