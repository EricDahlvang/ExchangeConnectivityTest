using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net;

namespace ExchangeConnectivityTest
{
    public class CredentialsTester
    {
        public static ExchangeServerSettings TestCredentials(string emailAddress, string password, bool trace = false)
        {
            try
            {
                var service = ConnectServiceWithFallback(emailAddress, password, trace);

                if (service == null)
                    return null;

                var version = service.GetExchangeVersionString();
                return new ExchangeServerSettings() { AutodiscoverUrl = service.Url, Version = version };
            }
            catch (Exception ex)
            {
                Console.Error.Write(ex);
                Console.WriteLine("-----");
            }

            return null;
        }

        static ExchangeService ConnectServiceWithFallback(string emailAddress, string password, bool trace)
        {
            // try Exchange 2013 first and then fallback if needed
            ExchangeVersion exchangeVersion = ExchangeVersion.Exchange2013_SP1;
            try
            {
                string url = null;
                while ((int)exchangeVersion >= 0)
                {
                    Console.WriteLine("Testing Exchange Version: " + exchangeVersion);

                    var service = ConnectService(emailAddress, password, url, exchangeVersion, trace);

                    url = service.Url.ToString();

                    try
                    {
                        Folder.Bind(service, WellKnownFolderName.Root);
                        return service;
                    }
                    catch (ServiceVersionException ex)
                    {
                        Console.Error.Write(ex);
                        Console.WriteLine("-----");
                        exchangeVersion = (ExchangeVersion)((int)exchangeVersion) - 1;
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.Write(e);
            }

            return null;
        }

        static ExchangeService ConnectService(string emailAddress, string password, string url, ExchangeVersion version = ExchangeVersion.Exchange2013_SP1, bool trace = false)
        {
            ImpersonatedUserId impersonatedUserId = null;
            NetworkCredential networkCredentials = null;

            if (networkCredentials == null)
                networkCredentials = new NetworkCredential(emailAddress, password);

            var service = new ExchangeService(version, TimeZoneInfo.Utc)
            {
                ImpersonatedUserId = impersonatedUserId,
                Credentials = networkCredentials,
                // SCP is slow and only applies to in-network Exchange 
                // see https://blogs.msdn.microsoft.com/webdav_101/2015/05/11/best-practices-ews-authentication-and-access-issues/
                EnableScpLookup = false,
                UserAgent = "MSBotFramework"
            };

            // set X-AnchorMailbox when impersonation is used
            // https://blogs.msdn.microsoft.com/webdav_101/2015/05/11/best-practices-ews-authentication-and-access-issues/
            // https://blogs.msdn.microsoft.com/mstehle/2013/07/25/managing-affinity-for-ews-impersonation-in-exchange-2013-and-exchange-online-w15/
            if (service.ImpersonatedUserId?.Id != null)
                service.HttpHeaders.Add("X-AnchorMailbox", service.ImpersonatedUserId.Id);

            if (trace)
            {
                service.TraceEnablePrettyPrinting = true;
                service.TraceFlags = TraceFlags.All;
                service.TraceListener = ConsoleTraceListener.Instance;
                service.TraceEnabled = true;
            }

            if (!string.IsNullOrEmpty(url))
            {
                service.Url = new Uri(url);
            }
            else
            {
                service.AutodiscoverUrl(emailAddress, RedirectionUrlValidationCallback);
            }

            return service;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            var redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            return string.Equals("https", redirectionUri.Scheme, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class ExchangeServerSettings
    {
        public string Version { get; set; }

        public Uri AutodiscoverUrl { get; set; }
    }

    class ConsoleTraceListener : ITraceListener
    {
        public static object _instanceLock = new object();
        static ITraceListener _listener;

        public static ITraceListener Instance
        {
            get
            {
                if (_listener == null)
                {
                    lock (_instanceLock)
                    {
                        if (_listener == null)
                            _listener = new ConsoleTraceListener();
                    }
                }
                return _listener;
            }
        }

        public void Trace(string traceType, string traceMessage)
        {
            Console.WriteLine($"Trace: {traceType} Message:{traceMessage}");
        }
    }

    public static class EWSExtensions
    {
        public static string GetExchangeVersionString(this ExchangeService service)
        {
            // ServerInfo isn't populated until we make a call to the server
            if (service.ServerInfo == null)
                Folder.Bind(service, WellKnownFolderName.Root);

            var v = service.ServerInfo;
            return $"{v.MajorVersion}.{v.MinorVersion}.{v.MajorBuildNumber}.{v.MinorBuildNumber}";
        }
    }
}
