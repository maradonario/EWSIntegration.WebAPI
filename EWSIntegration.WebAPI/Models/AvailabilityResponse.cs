using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace EWSIntegration.WebAPI.Models
{
    public class AvailabilityResponse
    {
        public ICollection<AvailabilityUser> AvailabilityResult { get; set; }

    }

    public class AvailabilityUser
    {
        public ICollection<TimeBlock> Availability { get; set; }
    }

    public class TimeBlock
    {
        public DateTime Start { get; set; }

        public DateTime End { get; set; }

        public LegacyFreeBusyStatus StatusEnum { get; set; }

        public string Status { get; set; }
    }
}
