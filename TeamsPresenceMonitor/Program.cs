using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace TeamsPresenceMonitor
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();

            var clientId = configuration["ClientId"];
            var tenantId = configuration["TenantId"];
            var scopes = configuration["Scopes"].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); ;

            var graphAuthProvider = new GraphAuthProvider(clientId, tenantId, scopes);

            var graphClient = new GraphServiceClient(
                "https://graph.microsoft.com/beta",
                new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    var tokens = await graphAuthProvider.GetGraphTokenAsync();

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", tokens.AccessToken);
                }));

            while (true)
            {
                try
                {
                    var presence = await graphClient.Me.Presence.Request().GetAsync();

                    Console.WriteLine($"{DateTime.Now.ToString("g")} - {presence.Availability} - {presence.Activity}");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"An error has occured: {e.Message}");
                }

                Thread.Sleep(TimeSpan.FromSeconds(10));
            }
        }
    }
}
