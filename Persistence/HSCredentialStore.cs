using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace HorizonScientific
{
    // DPAPI-encrypted credential store for the Spincraft Persistence assembly.
    //
    // All credentials are stored as key/value pairs in a single encrypted file:
    //   %ProgramData%\Spincraft\sc.cred
    //
    // The file is encrypted with DataProtectionScope.LocalMachine, meaning:
    //   - Any process running on the same machine can decrypt it.
    //   - Copying the file to a different machine will not decrypt.
    //   - Plaintext credentials are never written to disk.
    //
    // DEPLOYMENT — run CredentialManager.exe once on every machine where the
    // assembly is deployed, or call Set() directly:
    //
    //   HSCredentialStore.Set(HSCredentialStore.Keys.NoReplyLogin,    "noreply@standexetg.us");
    //   HSCredentialStore.Set(HSCredentialStore.Keys.NoReplyPassword, "...");

    public static class HSCredentialStore
    {
        // Well-known credential keys. Add new entries here as additional accounts are onboarded.
        public static class Keys
        {
            public const string SpincraftServiceAccountId = "SpincraftServiceAccountId";
            public const string SpincraftServicePassword = "SpincraftServicePassword";
            public const string NoReplyLogin    = "NoReplyLogin";
            public const string NoReplyPassword = "NoReplyPassword";
        }

        private static readonly string CREDENTIAL_FILE_PATH =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Spincraft", "sc.cred");

        private static readonly object _lock = new object();
        private static Dictionary<string, string> _cache;

        // Retrieve a single credential value at runtime.
        // Throws InvalidOperationException if the store has not been provisioned.
        // Throws KeyNotFoundException if the key does not exist in the store.
        public static string Get(string key)
        {
            lock (_lock)
            {
                EnsureLoaded();
                if (_cache.TryGetValue(key, out string value))
                    return value;

                throw new KeyNotFoundException(
                    $"Credential key '{key}' was not found in the store at: {CREDENTIAL_FILE_PATH}. " +
                    $"Call HSCredentialStore.Set(\"{key}\", value) to add it.");
            }
        }

        // Add or update a single credential, preserving all other stored values.
        // Safe to call multiple times; each call re-reads, merges, and re-writes the file.
        public static void Set(string key, string value)
        {
            lock (_lock)
            {
                Dictionary<string, string> credentials =
                    File.Exists(CREDENTIAL_FILE_PATH) ? LoadFromDisk() : new Dictionary<string, string>();

                credentials[key] = value;
                SaveToDisk(credentials);
                _cache = null; // invalidate so next Get re-reads from disk
            }
        }

        // Overwrite the entire credential file with the provided dictionary.
        // Use this to provision all credentials on a new machine in one call.
        public static void Provision(IDictionary<string, string> credentials)
        {
            lock (_lock)
            {
                SaveToDisk(new Dictionary<string, string>(credentials));
                _cache = null;
            }
        }

        // -------------------------------------------------------------------------

        private static void EnsureLoaded()
        {
            if (_cache != null)
                return;

            if (!File.Exists(CREDENTIAL_FILE_PATH))
                throw new InvalidOperationException(
                    "The HSCredentialStore has not been provisioned on this machine. " +
                    "Run CredentialManager.exe or call HSCredentialStore.Set(key, value) for each credential. " +
                    "Expected file location: " + CREDENTIAL_FILE_PATH);

            _cache = LoadFromDisk();
        }

        private static Dictionary<string, string> LoadFromDisk()
        {
            byte[] encryptedBytes = File.ReadAllBytes(CREDENTIAL_FILE_PATH);
            byte[] plainBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.LocalMachine);

            var result = new Dictionary<string, string>();
            foreach (string line in Encoding.UTF8.GetString(plainBytes).Split('\n'))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split on the first '=' only, so passwords containing '=' are handled correctly.
                int sep = line.IndexOf('=');
                if (sep < 1)
                    continue;

                result[line.Substring(0, sep).Trim()] = line.Substring(sep + 1).Trim('\r');
            }
            return result;
        }

        private static void SaveToDisk(Dictionary<string, string> credentials)
        {
            // key=value lines joined by '\n'. Values may contain '=' — only the first is a separator.
            string plainText = string.Join("\n", credentials.Select(kvp => kvp.Key + "=" + kvp.Value));
            byte[] plainBytes = Encoding.UTF8.GetBytes(plainText);
            byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.LocalMachine);

            string directory = Path.GetDirectoryName(CREDENTIAL_FILE_PATH);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            File.WriteAllBytes(CREDENTIAL_FILE_PATH, encryptedBytes);
        }
    }
}
