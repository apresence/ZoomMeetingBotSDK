using System;
using System.Collections.Specialized;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Text;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;

namespace ZoomMeetingBotSDK
{
    /// Class library for creating tokens for use with Zoom APIs. Its primary utility is enabling non-interative applications to call Zoom APIs; Specifically the Zoom Client SDK C# Wrapper.
    /// At the time this library was written, the information on how to generate these tokens was not available in any one place, and a reliable implementation in C# could not be found.
    /// More than a dozen URLs worth of scattered documentation and a solid day's worth of work was required to create it.
    /// More information on the Zoom Client SDK C# Wrapper:
    ///   https://marketplace.zoom.us/docs/sdk/native-sdks/windows/c-sharp-wrapper/
    public static class ZoomTokens
    {
        private static HttpClient _httpClient;
        private static string _zoomWebDomain;


        public static void Init(string ZoomWebDomain)            
        {
            _httpClient = new HttpClient();
            _zoomWebDomain = ZoomWebDomain;
        }

        private static string _CallZoomWebAPI(HttpMethod method, string requestURI, string authScheme, string authParam, NameValueCollection query)
        {
            using (var httpRequestMessage = new HttpRequestMessage(method, $"{_zoomWebDomain}{requestURI}?{query}"))
            {
                httpRequestMessage.Headers.Authorization = new AuthenticationHeaderValue(authScheme, authParam);
                var response = _httpClient.SendAsync(httpRequestMessage).Result;
                response.EnsureSuccessStatusCode();
                return response.Content.ReadAsStringAsync().Result;
            }
        }

        /// <summary>
        /// Wrapper class for functions used to create OAuth Tokens for use with Zoom APIs.
        /// </summary>
        public static class OAuth
        {
            /// <summary>
            /// Creates an OAuth token for Zoom Account Credentials using the "Server-to-Server OAuth" Zoom App.
            /// Allows the token to be generated without interactively prompting for credentials or requiring a URL redirect.
            /// Useful when creating unattended / non-interactive applications that interface with Zoom.
            /// Details:
            ///   https://marketplace.zoom.us/docs/guides/build/server-to-server-oauth-app/
            ///   https://devforum.zoom.us/t/switching-from-non-oauth-to-oauth-for-macos-sdk-application-with-active-users/67051
            ///   https://jenzushsu.medium.com/setting-up-server-to-server-s2s-oauth-to-test-zoom-apis-via-postman-32c9cd7a73
            /// </summary>
            /// <param name="accountID">The "Account ID" obtained from your Server-to-Server OAuth App's "App Credentials" page.</param>
            /// <param name="clientID">The "Client ID" obtained from your Server-to-Server OAuth App's "App Credentials" page.</param>
            /// <param name="clientSecret">The "Client secret" obtained from your Server-to-Server OAuth App's "App Credentials" page.</param>
            /// <returns>A string containing the generated OAuth Token.</returns>
            public static string CreateS2SToken(string accountID, string clientID, string clientSecret)
            {
                var q = HttpUtility.ParseQueryString(string.Empty);
                q["grant_type"] = "account_credentials";
                q["account_id"] = accountID;

                return _CallZoomWebAPI(
                    HttpMethod.Post,
                    "/oauth/token",
                    "Basic",
                    System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{clientID}:{clientSecret}")),
                    q
                );

                /*

                using (var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, $"{_zoomWebDomain}/oauth/token?{q}"))
                {
                    httpRequestMessage.Headers.Authorization = new AuthenticationHeaderValue("Basic", System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{clientID}:{clientSecret}")));
                    var response = _httpClient.SendAsync(httpRequestMessage).Result;
                    response.EnsureSuccessStatusCode();
                    return response.Content.ReadAsStringAsync().Result;
                }
                */
            }
        }

        public static class ZAK
        {
            /// <summary>
            /// Creates a Zoom Access Key (ZAK) using the given OAuth token.
            /// The ZAK can then be used to call Zoom APIs.
            /// Details:
            ///   https://marketplace.zoom.us/docs/api-reference/zoom-api/methods/#operation/userToken
            ///   https://devforum.zoom.us/t/generate-authorization-code-and-token-c-code/39197/2
            /// </summary>
            /// <param name="oAuthToken">An OAuth token generated by GetS2SOAuthToken() or similar.</param>
            /// <param name="userID">The Zoom User ID for which the token should be generated. Use "me" (the default) for user-level apps.</param>
            /// <param name="validityPeriod">How long the token should last.</param>
            /// <returns>A string containing the generated ZAK.</returns>
            public static string CreateToken(string oAuthToken, string userID = "me", TimeSpan? validityPeriod = null)
            {
                var validitySeconds = (validityPeriod == null) ? TimeSpan.FromHours(2).TotalSeconds : validityPeriod.Value.TotalSeconds;

                var q = HttpUtility.ParseQueryString(string.Empty);
                q["type"] = "zak";
                q["tti"] = $"{validitySeconds}";

                return _CallZoomWebAPI(
                    HttpMethod.Get,
                    $"/v2/users/{HttpUtility.UrlEncode(userID)}/token",
                    "Bearer",
                    oAuthToken,
                    q
                );

                /*
                using (var httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, $"{_zoomWebDomain}/v2/users/{HttpUtility.UrlEncode(userID)}/token?{q}"))
                {
                    httpRequestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", oAuthToken);
                    var response = _httpClient.SendAsync(httpRequestMessage).Result;
                    response.EnsureSuccessStatusCode();
                    return response.Content.ReadAsStringAsync().Result;
                }
                */
            }
        }

        /// <summary>
        /// Wrapper class for functions used to create Java Web Tokens for use with Zoom APIs.  At the time this library was written, the information on how to
        /// generate these tokens was not available in any one place, and a reliable implementation in C# could not be found.  This is an amalgamation of the
        /// following:
        ///   https://marketplace.zoom.us/docs/guides/auth/jwt
        ///   https://github.com/zoom/zoom-sdk-windows/blob/master/CHANGELOG.md#new-sdk-initialization-method-using-jwt-token
        ///   https://devforum.zoom.us/t/zoom-sdk-initialize-failed-when-running-the-sample-project/26511
        ///   https://devforum.zoom.us/t/how-to-create-jwt-token-using-rest-api-in-c/6620
        /// </summary>
        public static class JWT
        {
            /// <summary>
            /// Creates a Java Web Token for use with the Zoom Client SDK: https://marketplace.zoom.us/docs/sdk/native-sdks/introduction.
            /// </summary>
            /// <param name="sdkKey">The "SDK Key" obtained from your SDK App's "App Credentials" page.</param>
            /// <param name="sdkSecret">The "SDK Secret" obtained from your SDK App's "App Credentials" page.</param>
            /// <param name="accessTokenValidityInMinutes">How long the access token is valid for, in minutes.  Maximum value: 48 hours.</param>
            /// <param name="tokenValidityInMinutes">How long the token is valid for, in minutes.  Minimum value: 30 minutes.  Defaults to accessTokenValidityInMinutes.</param>
            /// <returns>A string containing the generated Java Web Token.</returns>
            public static string CreateClientSDKToken(string sdkKey, string sdkSecret, int accessTokenValidityInMinutes = 120, int tokenValidityInMinutes = -1)
            {
                DateTime now = DateTime.UtcNow;

                if (tokenValidityInMinutes <= 0)
                {
                    tokenValidityInMinutes = accessTokenValidityInMinutes;
                }

                int tsNow = (int)(now - new DateTime(1970, 1, 1)).TotalSeconds;
                int tsAccessExp = (int)(now.AddMinutes(accessTokenValidityInMinutes) - new DateTime(1970, 1, 1)).TotalSeconds;
                int tsTokenExp = (int)(now.AddMinutes(tokenValidityInMinutes) - new DateTime(1970, 1, 1)).TotalSeconds;

                return CreateToken(sdkSecret, new JwtPayload
                {
                    { "appKey", sdkKey },
                    { "iat", tsNow },
                    { "exp", tsAccessExp },
                    { "tokenExp", tsTokenExp },
                });
            }

            /// <summary>
            /// Creates a Java Web Token for use with the Zoom API: https://marketplace.zoom.us/docs/api-reference/zoom-api.
            /// </summary>
            /// <param name="apiKey">The "SDK Key" obtained from your SDK App's "App Credentials" page.</param>
            /// <param name="apiSecret">The "SDK Secret" obtained from your SDK App's "App Credentials" page.</param>
            /// <param name="tokenValidityInMinutes">How long the token is valid for, in minutes.  Minimum value: 30 minutes.</param>
            /// <returns>A string containing the generated Java Web Token.</returns>
            public static string CreateAPIToken(string apiKey, string apiSecret, int tokenValidityInMinutes = 120)
            {
                DateTime now = DateTime.UtcNow;

                int tsNow = (int)(now - new DateTime(1970, 1, 1)).TotalSeconds;
                int tokenExp = (int)(now.AddMinutes(tokenValidityInMinutes) - new DateTime(1970, 1, 1)).TotalSeconds;

                return CreateToken(apiSecret, new JwtPayload
                {
                    { "appKey", apiKey },
                    { "iat", tsNow },
                    { "tokenExp", tokenExp },
                });
            }

            /// <summary>
            /// Creates a Java Web Token with the given secret and payload.  This should not be called directly unless you intend to hand-craft a JWT's payload.
            /// </summary>
            /// <param name="secret">Encryption key.</param>
            /// <param name="payload">JWT Payload.</param>
            /// <returns>A string containing the generated Java Web Token.</returns>
            public static string CreateToken(string secret, JwtPayload payload)
            {
                // Create Security key using private key above:
                // note that latest version of JWT using Microsoft namespace instead of System
                var securityKey = new Microsoft.IdentityModel.Tokens.SymmetricSecurityKey(Encoding.UTF8.GetBytes(secret));

                // Also note that securityKey length should be >256b
                // so you have to make sure that your private key has a proper length
                var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);

                // Finally create a Token
                var header = new JwtHeader(credentials);

                var secToken = new JwtSecurityToken(header, payload);
                var handler = new JwtSecurityTokenHandler();

                // Token to String so you can use it in your client
                return handler.WriteToken(secToken);
            }
        }
    }
}
