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
        public ICollection<Interview> Appointments { get; set; }
    }

    /// <summary>
    /// Interview
    /// </summary>
    public class Interview
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

        /// <summary>
        /// Gets or sets the attendees.
        /// </summary>
        /// <value>
        /// The attendees.
        /// </value>
        public ICollection<RequiredAttendees> Attendees { get; set;}
    }

    /// <summary>
    /// 
    /// </summary>
    public class RequiredAttendees
    {
        /// <summary>
        /// Gets or sets the email.
        /// </summary>
        /// <value>
        /// The email.
        /// </value>
        public string Email { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the response.
        /// </summary>
        /// <value>
        /// The response.
        /// </value>
        public string Response { get; set; }
    }

}
