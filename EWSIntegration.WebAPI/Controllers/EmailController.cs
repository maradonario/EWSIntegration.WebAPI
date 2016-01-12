using EWSIntegration.WebAPI.Models;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace EWSIntegration.WebAPI.Controllers
{
    public class EmailController : ApiController
    {

        [HttpPost]
        public IHttpActionResult Send(SendEmailRequest request)
        {
            // Send Email
            var email = new EmailMessage(ExchangeServer.Open());

            email.ToRecipients.AddRange(request.Recipients);

            email.Subject = request.Subject;
            email.Body = new MessageBody(request.Body);
            email.Send();
            return Ok(new SendEmailResponse { Recipients = request.Recipients });
        }
    }
}
