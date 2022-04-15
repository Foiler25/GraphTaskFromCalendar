using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Extensions.Configuration;
using RestSharp;
using RestSharp.Authenticators;
using System.Text.Json;


namespace GraphTaskFromCalendar {
    public class DeviceCodeFlowAuthorizationProvider : IAuthenticationProvider
    {
        private readonly IPublicClientApplication _application;
        private readonly List<string> _scopes;        
        //private string _authToken = Program.LoadAppSettings.config["accesstoken"];
        private string _grantType;
        private string _clientId;
        private string _clientSecret;
        private string _resource;
        private string _username;
        private string _password;
        private string _website;
        private string _authToken;
        public DeviceCodeFlowAuthorizationProvider(IPublicClientApplication application, List<string> scopes, List<string> login) 
        {
            _grantType = login[0];
            _clientId = login[1];
            _clientSecret = login[2];
            _resource = login[3];
            _username = login[4];
            _password = login[5];
            _website = login[6];
            _application = application;
            _scopes = scopes;
        }
        public class ResponseObj
        {
            public string token_type { get; set; }
            public string scope { get; set; }
            public string expires_in { get; set; }
            public string exp_expires_in { get; set; }
            public string expires_on { get; set; }
            public string not_before { get; set; }
            public string resource { get; set; }
            public string access_token { get; set; }
            public string refresh_token { get; set; }
        }
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {

            if(string.IsNullOrEmpty(_authToken))
            {
                var client = new RestClient(_website);
                var call = new RestRequest();
                call.AlwaysMultipartFormData = true;
                call.AddParameter("grant_type", _grantType);
                call.AddParameter("client_id", _clientId);
                call.AddParameter("client_secret", _clientSecret);
                call.AddParameter("resource", _resource);
                call.AddParameter("username", _username);
                call.AddParameter("password", _password);
                var response = client.PostAsync(call);
                //Console.WriteLine(response.Result.Content);

                ResponseObj deserializedResult = JsonSerializer.Deserialize<ResponseObj>(response.Result.Content);
                _authToken = deserializedResult.access_token;
                //Console.WriteLine(_authToken);

                if(string.IsNullOrEmpty(_authToken))
                {
                    var result = await _application.AcquireTokenWithDeviceCode(_scopes, callback => {
                    Console.WriteLine(callback.Message);
                    return Task.FromResult(0);
                    }).ExecuteAsync();
                }
                
            }
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", _authToken);
        }
    }
}