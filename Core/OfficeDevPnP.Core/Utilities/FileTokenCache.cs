#if !NETSTANDARD2_0
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Class deals with Token Caching of file
    /// </summary>
    public class FileTokenCache : TokenCache
    {
        public string CacheFilePath;
        private static readonly object FileLock = new object();

        /// <summary>
        /// Constructor for FileTokenCache class
        /// </summary>
        /// <param name="filePath">Path of the file</param>
        public FileTokenCache(string filePath = @".\TokenCache.dat")
        {
            CacheFilePath = filePath;
            this.AfterAccess = AfterAccessNotification;
            this.BeforeAccess = BeforeAccessNotification;
            lock (FileLock)
            {
                this.Deserialize(File.Exists(CacheFilePath) ?
                    ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath),
                                            null,
                                            DataProtectionScope.CurrentUser)
                    : null);
            }
        }

        /// <summary>
        /// Clears the Cached file path
        /// </summary>
        public override void Clear()
        {
            base.Clear();
            File.Delete(CacheFilePath);
        }

        void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                this.Deserialize(File.Exists(CacheFilePath) ?
                    ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath),
                                            null,
                                            DataProtectionScope.CurrentUser)
                    : null);
            }
        }

        void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            if (this.HasStateChanged)
            {
                lock (FileLock)
                {
                    File.WriteAllBytes(CacheFilePath,
                        ProtectedData.Protect(this.Serialize(),
                                                null,
                                                DataProtectionScope.CurrentUser));
                    this.HasStateChanged = false;
                }
            }
        }
    }
}
#endif