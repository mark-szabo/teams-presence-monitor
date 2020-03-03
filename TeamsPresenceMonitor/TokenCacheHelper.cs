using Microsoft.Identity.Client;
using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;

namespace TeamsPresenceMonitor
{
    public static class TokenCacheHelper
    {
        public static void EnableSerialization(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
        }

        /// <summary>
        /// Path to the token cache
        /// </summary>
        public static readonly string CacheFilePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\msalcache.bin3";

        private static readonly object FileLock = new object();

        private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                byte[] state = null;

                if (File.Exists(CacheFilePath))
                {
                    var encryptedState = File.ReadAllBytes(CacheFilePath);
                    state = ProtectedData.Unprotect(encryptedState, null, DataProtectionScope.CurrentUser);
                }

                args.TokenCache.DeserializeMsalV3(state);
            }
        }

        private static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                lock (FileLock)
                {
                    // reflect changesgs in the persistent store
                    var serializedState = args.TokenCache.SerializeMsalV3();
                    var encryptedState = ProtectedData.Protect(serializedState, null, DataProtectionScope.CurrentUser);
                    File.WriteAllBytes(CacheFilePath, encryptedState);
                }
            }
        }
    }
}
