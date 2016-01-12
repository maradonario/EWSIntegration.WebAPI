using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EWSIntegration.WebAPI.Controllers;
using System.Web.Http.Results;
using EWSIntegration.WebAPI.Models;

namespace EWSIntegration.WebAPI.Tests.Controllers
{
    /// <summary>
    /// Summary description for AppointmentsControllerTests
    /// </summary>
    [TestClass]
    public class AppointmentsControllerTests
    {
        public AppointmentsControllerTests()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void GetAppointment()
        {
            var controller = new AppointmentsController();

            var result = controller.Get() as  OkNegotiatedContentResult<GetAppointmentsResponse>;

            Assert.IsNotNull(result.Content.Appointments);
        }

        [TestMethod]
        public void CreateAppointment()
        {
            var controller = new AppointmentsController();

            var start = new DateTime(2016, 1, 14, 13, 0, 0);
            var end = start.AddHours(2);
            var request = new CreateAppointmentRequest
            {
                Body = "Appointment Created from Unit Test",
                Location = "Cconference Room",
                Subject = "Appointment Unit Test",
                Start = start.ToString(),
                End = end.ToString()
            };

            var result = controller.Create(request) as OkNegotiatedContentResult<CreateAppointmentResponse>;

            Assert.IsNotNull(result.Content.AppointId);
            Assert.IsNotNull(result.Content.Message);
        }
    }
}
