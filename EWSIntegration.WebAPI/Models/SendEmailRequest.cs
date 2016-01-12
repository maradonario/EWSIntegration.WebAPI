using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSIntegration.WebAPI.Models
{
    /// <summary>
    /// SendEmailRequest
    /// </summary>
    public class SendEmailRequest
    {

        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        /// <value>
        /// The subject.
        /// </value>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets the body.
        /// </summary>
        /// <value>
        /// The body.
        /// </value>
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets the recipients.
        /// </summary>
        /// <value>
        /// The recipients.
        /// </value>
        public ICollection<string> Recipients { get; set; }

        /// <summary>
        /// Gets or sets the file attachments.
        /// </summary>
        /// <value>
        /// The file attachments.
        /// </value>
        public ICollection<string> FileAttachments { get; set; }
    }
}
