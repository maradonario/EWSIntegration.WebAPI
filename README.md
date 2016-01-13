# EWSIntegration.WebAPI
Test Web API for EWS Managed API Integration

To use this Web API locally, modify the following web.config values appropiately:

    <!--These are the web service settgins for the service account-->
    <add key="SERVICE_ACCOUNT_EMAIL" value="email@email.com"/>
    <add key="SERVICE_ACCOUNT_PASS" value="password"/>
    <add key="SERVICE_URL" value="https://outlook.office365.com/EWS/Exchange.asmx"/>
    <add key="URL_AUTODISCOVER" value="false"/>
    
    SERVICE_ACCOUNT_EMAIL - email of the service account
    SERVICE_ACCOUNT_PASS - service account password
    SERVICE_URL - service URL
    URL_AUTODISCOVER - setting this to false, the ExchangeService will be initiated using the url provided in SERVICE_URL.
    
    To test connection, also set the above settings apporpiately in the Test Poject and run the unit tests.
