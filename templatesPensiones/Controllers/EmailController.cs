using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using templatesPensiones.Model;
using System.Net.Mail;

namespace templatesPensiones.Controllers
{
    public class EmailController : Controller
    {
        // GET: Email
        public ActionResult correoindex()
        {
            return View();
        }

        [HttpPost]
        public ActionResult correoindex(email em)
        {

            string to = em.To;
            string subject = em.Subject;
            string bodey = em.Body;
            MailMessage mm = new MailMessage();
            mm.To.Add(to);
            mm.Subject = subject;
            mm.Body = bodey;
            mm.From = new MailAddress("oficinapensionesoaxaca@gmail.com");   
          //  mm.From = new MailAddress("tecnologias@pensionesoaxaca.com");    //empresarial

            mm.IsBodyHtml = false;
            SmtpClient smtp = new SmtpClient("smtp.gmail.com");
            smtp.Port = 587;
            smtp.UseDefaultCredentials = true;
            smtp.EnableSsl = true;
          //  smtp.Credentials = new System.Net.NetworkCredential("tecnologias@pensionesoaxaca.com", "admindie");
           // permisos auth terceros
           smtp.Credentials = new System.Net.NetworkCredential("oficinapensionesoaxaca@gmail.com", "admindie");
            smtp.Send(mm);
            ViewBag.message = "Email fue enviado a " + em.To + " , con éxito";
            return View();
        }
    }
}