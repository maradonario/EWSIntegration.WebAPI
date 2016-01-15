using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Web.Http;
using EWSIntegration.WebAPI.Models;
namespace EWSIntegration.WebAPI.Controllers
{
    /// <summary>
    /// AppointmentsController. View or Create Appointments
    /// </summary>
    public class AppointmentsController : ApiController
    {
        /// <summary>
        /// Availabilities the specified request.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns></returns>
        [HttpPost]
        [Route("api/appointments/availability")]
        public IHttpActionResult Availability(AvailabilityRequest request)
        {
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();

            foreach (var user in request.Users)
            {
                attendees.Add(new AttendeeInfo()
                {
                    SmtpAddress = user,
                    AttendeeType = MeetingAttendeeType.Required
                });
            }

            // Specify availability options.
            AvailabilityOptions myOptions = new AvailabilityOptions();
            myOptions.MeetingDuration = request.DurationMinutes;
            myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

            // Return a set of free/busy times.
            var service = ExchangeServer.Open();

            var startTime = DateTime.Parse(request.Start);
            var endTime = DateTime.Parse(request.End);
            GetUserAvailabilityResults freeBusyResults = service.GetUserAvailability(attendees,
                                                                                 new TimeWindow(startTime, endTime),
                                                                                     AvailabilityData.FreeBusy,
                                                                                     myOptions);

            var response = new AvailabilityResponse
            {
                AvailabilityResult = new List<AvailabilityUser>()
            };


            foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
            {
                var user = new AvailabilityUser();
                var avail = new List<TimeBlock>();

                foreach (CalendarEvent calendarItem in availability.CalendarEvents)
                {
                    var block = new TimeBlock
                    {
                        Start = calendarItem.StartTime,
                        End = calendarItem.EndTime,
                        StatusEnum = calendarItem.FreeBusyStatus,
                        Status = calendarItem.FreeBusyStatus.ToString()
                    };

                    avail.Add(block);
                }
                user.Availability = avail;
                response.AvailabilityResult.Add(user);
            }

            return Ok(response);
        }
        /// <summary>
        /// Gets the appointments starting now to 30 days from now.
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        [Route("api/appointments/details")]
        public IHttpActionResult GetDetails(GetAppointmentsRequest request)
        {
            var startDate = DateTime.Parse(request.Start);
            var endDate = DateTime.Parse(request.End);
            // Assuming 8 max per day * 5 (days/week) * 4 (weeks/month)
            const int NUM_APPTS = 160;
            var calendar = CalendarFolder.Bind(ExchangeServer.Open(), WellKnownFolderName.Calendar, new PropertySet());
            var cView = new CalendarView(startDate, endDate, NUM_APPTS);
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.TimeZone);



            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            var response = new GetAppointmentsResponse();
            var list = new List<Interview>();
            foreach(var app in appointments)
            {
                var appointment = Appointment.Bind(ExchangeServer.Open(), app.Id);

                var attendees = new List<RequiredAttendees>();
                foreach (var required in appointment.RequiredAttendees)
                {
                    attendees.Add(new RequiredAttendees
                    {
                        Name = required.Name,
                        Email = required.Address,
                        Response = required.ResponseType.ToString()
                    });


                }

                list.Add(new Interview { Start = app.Start, End = app.End, TimeZone = app.TimeZone, Attendees = attendees, Subject = app.Subject});

            }
            response.Appointments = list;
            return Ok(response);
        }

        /// <summary>
        /// Creates the specified Interview represented by the request.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns></returns>
        [HttpPost]
        public IHttpActionResult Create(CreateAppointmentRequest request)
        {
            var service = ExchangeServer.Open();
            var appointment = new Appointment(service);

            // Set the properties on the appointment object to create the appointment.
            appointment.Subject = request.Subject;
            appointment.Body = request.Body;
            appointment.Start = DateTime.Parse(request.Start);
            //appointment.StartTimeZone = TimeZoneInfo.Local;
            appointment.End = DateTime.Parse(request.End);
            //appointment.EndTimeZone = TimeZoneInfo.Local;
            appointment.Location = request.Location;
            //appointment.ReminderDueBy = DateTime.Now;

            foreach(var email in request.Recipients)
            {
                appointment.RequiredAttendees.Add(email);
            }

            // Save the appointment to your calendar.
            appointment.Save(SendInvitationsMode.SendOnlyToAll);

            // Verify that the appointment was created by using the appointment's item ID.
            Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));

            var response = new CreateAppointmentResponse
            {
                Message = "Appointment created: " + item.Subject,
                AppointId = appointment.Id.ToString()
            };

            return Ok(response);
        }
    }
}
