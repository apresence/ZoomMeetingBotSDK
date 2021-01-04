namespace ZoomMeetingBotSDK
{
    using System;
    using System.Security.Cryptography;
    using System.Text;

    /// <summary>
    /// Provides a facility to encrypt/decrypt a string using C#'s ProtectedData class which depends on Windows' native DPAPI.
    /// The encyrption/decryption is only good for the current Windows machine.  More info here:
    /// https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.protecteddata?redirectedfrom=MSDN&view=netframework-4.8
    /// </summary>  
    public class ProtectedString
    {
        private static int ENTROPY_LEN = 20;

        public static string Protect(string value, DataProtectionScope scope = DataProtectionScope.LocalMachine)
        {
            // Generate ENTROPY_LEN bytes of random data
            byte[] entropy = new byte[ENTROPY_LEN];
            using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(entropy);
            }

            // Convert the string value to UTF8 bytes
            byte[] valueBytes = Encoding.UTF8.GetBytes(value);

            // Protect the string value bytes, mixing in the random entropy
            byte[] valueEncryptedBytes = ProtectedData.Protect(valueBytes, entropy, scope);

            // Save the random entropy and protected value bytes in a single byte array
            byte[] protectedData = new byte[entropy.Length + valueEncryptedBytes.Length];
            Buffer.BlockCopy(entropy, 0, protectedData, 0, entropy.Length);
            Buffer.BlockCopy(valueEncryptedBytes, 0, protectedData, entropy.Length, valueEncryptedBytes.Length);

            // Convert byte array to Base64 string
            string ret = Convert.ToBase64String(protectedData);

            // Zero out temporary arrays
            Array.Clear(protectedData, 0, protectedData.Length);
            Array.Clear(valueEncryptedBytes, 0, valueEncryptedBytes.Length);
            Array.Clear(valueBytes, 0, valueBytes.Length);
            Array.Clear(entropy, 0, entropy.Length);

            return ret;
        }

        public static string Unprotect(string value, DataProtectionScope scope = DataProtectionScope.LocalMachine)
        {
            // Convert Base64 string to byte array
            byte[] protectedData = Convert.FromBase64String(value);

            // Extract entropy and protected value in bytes
            byte[] entropy = new byte[ENTROPY_LEN];
            byte[] valueEncryptedBytes = new byte[protectedData.Length - entropy.Length];
            Buffer.BlockCopy(protectedData, 0, entropy, 0, entropy.Length);
            Buffer.BlockCopy(protectedData, ENTROPY_LEN, valueEncryptedBytes, 0, valueEncryptedBytes.Length);

            // Unprotect the string value bytes
            byte[] valueBytes = ProtectedData.Unprotect(valueEncryptedBytes, entropy, scope);

            // Convert unprotected bytes to UTF8 string
            string ret = Encoding.UTF8.GetString(valueBytes);

            // Zero out temporary arrays
            Array.Clear(valueBytes, 0, valueBytes.Length);
            Array.Clear(valueEncryptedBytes, 0, valueEncryptedBytes.Length);
            Array.Clear(entropy, 0, entropy.Length);
            Array.Clear(protectedData, 0, protectedData.Length);

            // Return string
            return ret;
        }
    }
}
