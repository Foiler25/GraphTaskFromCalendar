using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;



namespace GraphTaskFromCalendar
{
    public class Program
    {
        private static GraphServiceClient _graphServiceClient;
        private static HttpClient _httpClient;

        public static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]) ||
                    string.IsNullOrEmpty(config["domain"]) ||
                    string.IsNullOrEmpty(config["username"]) ||
                    string.IsNullOrEmpty(config["password"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }
        
        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var redirectUri = config["redirectUri"];
            var grantType = config["grant_type"];
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var resource = config["resource"];
            var username = config["username"];
            var password = config["password"];
            var website = config["microsoftLogin"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            List<string> login = new List<string>();
            login.Add(grantType);
            login.Add(clientId);
            login.Add(clientSecret);
            login.Add(resource);
            login.Add(username);
            login.Add(password);
            login.Add(website);

            var pca = PublicClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .Build();
            return new DeviceCodeFlowAuthorizationProvider(pca, scopes, login);
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }
        static void Main(string[] args)
        {
            // Load appsettings.json
            var config = LoadAppSettings();
            if (null == config)
            {
                Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                return;
            }
        
            var GraphServiceClient = GetAuthenticatedGraphClient(config);
            var plannerHelper = new PlannerHelper(GraphServiceClient);
            plannerHelper.PlannerHelperCall().GetAwaiter().GetResult();



        }

    }
    
}
