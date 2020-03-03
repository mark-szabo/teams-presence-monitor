﻿using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace TeamsPresenceMonitor
{
    internal class GraphAuthProvider
    {
        private readonly string[] _scopes;
        private readonly IPublicClientApplication _app;

        public GraphAuthProvider(string clientId, string tenantId, string[] scopes)
        {
            _scopes = scopes;

            var authority = "https://login.microsoftonline.com/" + tenantId;

            _app = PublicClientApplicationBuilder
                    .Create(clientId)
                    .WithAuthority(authority)
                    .WithDefaultRedirectUri()
                    .Build();

            TokenCacheHelper.EnableSerialization(_app.UserTokenCache);
        }

        public async Task<AuthenticationResult> GetGraphTokenAsync()
        {

            var accounts = await _app.GetAccountsAsync();

            // All AcquireToken* methods store the tokens in the cache, so check the cache first
            try
            {
                return await _app.AcquireTokenSilent(_scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // No token found in the cache or AAD insists that a form interactive auth is required (e.g. the tenant admin turned on MFA)
                // If you want to provide a more complex user experience, check out ex.Classification 

                return await AcquireByDeviceCodeAsync(_app);
            }
        }

        private async Task<AuthenticationResult> AcquireByDeviceCodeAsync(IPublicClientApplication pca)
        {
            try
            {
                var result = await pca.AcquireTokenWithDeviceCode(_scopes,
                    deviceCodeResult =>
                    {
                        // This will print the message on the console which tells the user where to go sign-in using 
                        // a separate browser and the code to enter once they sign in.
                        // The AcquireTokenWithDeviceCode() method will poll the server after firing this
                        // device code callback to look for the successful login of the user via that browser.
                        // This background polling (whose interval and timeout data is also provided as fields in the 
                        // deviceCodeCallback class) will occur until:
                        // * The user has successfully logged in via browser and entered the proper code
                        // * The timeout specified by the server for the lifetime of this code (typically ~15 minutes) has been reached
                        // * The developing application calls the Cancel() method on a CancellationToken sent into the method.
                        //   If this occurs, an OperationCanceledException will be thrown (see catch below for more details).
                        Console.WriteLine(deviceCodeResult.Message);
                        return Task.FromResult(0);
                    }).ExecuteAsync();

                Console.WriteLine(result.Account.Username);
                return result;
            }
            // TODO: handle or throw all these exceptions
            catch (MsalServiceException ex)
            {
                // Kind of errors you could have (in ex.Message)

                // AADSTS50059: No tenant-identifying information found in either the request or implied by any provided credentials.
                // Mitigation: as explained in the message from Azure AD, the authoriy needs to be tenanted. you have probably created
                // your public client application with the following authorities:
                // https://login.microsoftonline.com/common or https://login.microsoftonline.com/organizations

                // AADSTS90133: Device Code flow is not supported under /common or /consumers endpoint.
                // Mitigation: as explained in the message from Azure AD, the authority needs to be tenanted

                // AADSTS90002: Tenant <tenantId or domain you used in the authority> not found. This may happen if there are 
                // no active subscriptions for the tenant. Check with your subscription administrator.
                // Mitigation: if you have an active subscription for the tenant this might be that you have a typo in the 
                // tenantId (GUID) or tenant domain name.
                throw ex;
            }
            catch (OperationCanceledException ex)
            {
                // If you use a CancellationToken, and call the Cancel() method on it, then this *may* be triggered
                // to indicate that the operation was cancelled. 
                // See https://docs.microsoft.com/en-us/dotnet/standard/threading/cancellation-in-managed-threads 
                // for more detailed information on how C# supports cancellation in managed threads.
                throw ex;
            }
            catch (MsalClientException ex)
            {
                // Possible cause - verification code expired before contacting the server
                // This exception will occur if the user does not manage to sign-in before a time out (15 mins) and the
                // call to `AcquireTokenWithDeviceCode` is not cancelled in between
                throw ex;
            }
        }
    }


}
