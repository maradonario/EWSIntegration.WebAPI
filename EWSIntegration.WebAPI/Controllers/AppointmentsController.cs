using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Web.Http;
using AppointmentWebApi = EWSIntegration.WebAPI.Models.Appointment;
using ResponseWebApi = EWSIntegration.WebAPI.Models.GetAppointmentsResponse;
using CreateAppointmentRequest = EWSIntegration.WebAPI.Models.CreateAppointmentRequest;
using CreateAppointmentResponse = EWSIntegration.WebAPI.Models.CreateAppointmentResponse;
namespace EWSIntegration.WebAPI.Controllers
{
    /// <summary>
    /// AppointmentsController. View or Create Appointments
    /// </summary>
    public class AppointmentsController : ApiController
    {
        /// <summary>
        /// Gets the appointments starting now to 30 days from now.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IHttpActionResult Get()
        {
            var startDate = DateTime.Now;
            var endDate = startDate.AddDays(30);
            // Assuming 8 max per day * 5 (days/week) * 4 (weeks/month)
            const int NUM_APPTS = 160;

            var calendar = CalendarFolder.Bind(ExchangeServer.Open(), WellKnownFolderName.Calendar, new PropertySet());
            var cView = new CalendarView(startDate, endDate, NUM_APPTS);
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.TimeZone);

            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            var response = new ResponseWebApi();
            var list = new List<AppointmentWebApi>();
            foreach(var app in appointments)
            {
                list.Add(new AppointmentWebApi { Start = app.Start.ToUniversalTime().ToString(), End = String.Format("{0:g}", app.End), TimeZone = app.TimeZone});
            }
            response.Appointments = list;
            return Ok(response);
        }

        /// <summary>
        /// Creates the specified Appointment represented by the request.
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
