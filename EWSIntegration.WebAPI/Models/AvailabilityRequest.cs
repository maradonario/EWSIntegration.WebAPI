using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSIntegration.WebAPI.Models
{
    /// <summary>
    /// 
    /// </summary>
    public class AvailabilityRequest
    {
        /// <summary>
        /// Gets or sets the users.
        /// </summary>
        /// <value>
        /// The users.
        /// </value>
        public ICollection<string> Users { get; set; }

        /// <summary>
        /// Gets or sets the duration minutes.
        /// </summary>
        /// <value>
        /// The duration minutes.
        /// </value>
        public int DurationMinutes { get; set; }

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


    }
}
