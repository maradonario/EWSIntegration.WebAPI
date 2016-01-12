using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSIntegration.WebAPI.Models
{
    /// <summary>
    /// CreateAppointmentResponse
    /// </summary>
    public class CreateAppointmentResponse
    {
        /// <summary>
        /// Gets or sets the appoint identifier.
        /// </summary>
        /// <value>
        /// The appoint identifier.
        /// </value>
        public string AppointId { get; set; }

        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        /// <value>
        /// The message.
        /// </value>
        public string Message { get; set; }
    }
}
