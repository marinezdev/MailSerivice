using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

namespace WebServiceCorreo
{
    /// <summary>
    /// Descripción breve de Correo
    /// </summary>
    [WebService(Namespace = "http://www.asae.com.mx")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente. 
    // [System.Web.Script.Services.ScriptService]
    public class Correo : System.Web.Services.WebService
    {

        [WebMethod]
        public string CorreoMetPrivado(string host, int puerto, string usuario, string contra, string Email, string from, string Subject, string body)
        {
            string cadena = "";

            Correos correo = new Correos();
            if (correo.CorreoMetPrivadoEnvia(host, puerto, usuario, contra, Email, from, Subject, body) == "ok")
            {
                cadena = "Correo enviado";
            }
            else
            {
                cadena = "Error de envio";
            }
            
            return cadena;
        }

        [WebMethod]
        public string CorreoAsaeTickets(string host, int puerto, string usuario, string contra, string Email, string EmailCopia, string from, string Subject, string body)
        {
            string cadena = "";

            Correos correo = new Correos();
            if (correo.CorreoAsaeTicketsEnvia(host, puerto, usuario, contra, Email, EmailCopia, from, Subject, body) == "ok")
            {
                cadena = "Correo enviado";
            }
            else
            {
                cadena = "Error de envio";
            }

            return cadena;
        }

        [WebMethod]
        public string CorreoAsaeCRM_AgendaEvento(string host, int puerto, string usuario, string contra, string Email, string EmailCopia, string from , string Subject, string body, string archivoNombre, string EventoTitulo, DateTime EventoFechaI, DateTime EventoFechaF, string EventoOrganizaNombre, string EventoOrganizaMail, string EventoAsunto, string EventoDescripcion, string EventoUbicacion)
        {
            string resultado = "";
            Correos correo = new Correos();
            resultado = correo.CorreoAsaeCRM_AgendaEvento(host, puerto, usuario, contra, Email, EmailCopia, from, Subject, body, archivoNombre, EventoTitulo, EventoFechaI, EventoFechaF, EventoOrganizaNombre, EventoOrganizaMail, EventoAsunto, EventoDescripcion, EventoUbicacion);
            if (resultado == "ok")
            {
                resultado = "Correo enviado";
            }
            else
            {
                resultado = "Error de envio. \t" + resultado;
            }
            return resultado;
        }
    }
}
