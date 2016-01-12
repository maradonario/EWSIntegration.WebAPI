using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;

namespace EWSIntegration.WebAPI
{
    public static class ExchangeServer
    {
        private const string SERVICE_ACCT_EMAIL = "SERVICE_ACCOUNT_EMAIL";
        private const string SERVICE_ACCT_PASSWORD = "SERVICE_ACCOUNT_PASS";
        private const string SERVICE_URL = "SERVICE_URL";
        private const string USE_AUTODISCOVER = "URL_AUTODISCOVER";

        public static ExchangeService Open()
        {
            var email = ConfigurationManager.AppSettings[SERVICE_ACCT_EMAIL];
            var pass = ConfigurationManager.AppSettings[SERVICE_ACCT_PASSWORD];
            var autoDiscover = ConfigurationManager.AppSettings[USE_AUTODISCOVER];
            var autoDiscoverBool = Boolean.Parse(autoDiscover);

            return Open(email, pass, autoDiscoverBool);

        }

        public static ExchangeService Open(string email, string password, bool autodiscover)
        {
            var url = ConfigurationManager.AppSettings[SERVICE_URL];

            var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            service.Credentials = new WebCredentials(email, password);
            service.UseDefaultCredentials = false;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;



            if (autodiscover)
            {
                service.AutodiscoverUrl(email, RedirectionUrlValidationCallback);
            }
            else
            {
                service.Url = new Uri(url);
            }
            return service;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
