using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSIntegration.WebAPI.Models
{
    /// <summary>
    /// GetAppointmentsResponse
    /// </summary>
    public class GetAppointmentsResponse
    {
        /// <summary>
        /// Gets or sets the appointments.
        /// </summary>
        /// <value>
        /// The appointments.
        /// </value>
        public ICollection<Appointment> Appointments { get; set; }
    }

    /// <summary>
    /// Appointment
    /// </summary>
    public class Appointment
    {
        /// <summary>
        /// Gets or sets the start.
        /// </summary>
        /// <value>
        /// The start.
        /// </value>
        public string Start { get; set; }

        /// <summary>
        /// Gets or sets the end.
        /// </summary>
        /// <value>
        /// The end.
        /// </value>
        public string End { get; set; }

        /// <summary>
        /// Gets or sets the time zone.
        /// </summary>
        /// <value>
        /// The time zone.
        /// </value>
        public string TimeZone { get; set; }
    }
}
