using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IdentityModel.Selectors;
using System.IdentityModel.Tokens;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.ServiceModel;
using System.Web;
using System.Web.Configuration;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.IdentityModel.SecurityTokenService;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using AudienceRestriction = Microsoft.IdentityModel.Tokens.AudienceRestriction;
using AudienceUriValidationFailedException = Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException;
using SecurityTokenHandlerConfiguration = Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration;
using X509SigningCredentials = Microsoft.IdentityModel.SecurityTokenService.X509SigningCredentials;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token {
	public static partial class TokenHelper {
		#region フィールド

		/// <summary>
		/// SharePoint principal.
		/// </summary>
		public const string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

		/// <summary>
		/// Lifetime of HighTrust access token, 12 hours.
		/// </summary>
		public static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

		//
		// Configuration Constants
		//
		private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
		private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
		private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
		private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
		public const string S2SProtocol = "OAuth2";
		public const string DelegationIssuance = "DelegationIssuance1.0";
		private const string NameIdentifierClaimType = JsonWebTokenConstants.ReservedClaims.NameIdentifier;
		private const string TrustedForImpersonationClaimType = "trustedfordelegation";
		private const string ActorTokenClaimType = JsonWebTokenConstants.ReservedClaims.ActorToken;

		//
		// Environment Constants
		//
		private static string GlobalEndPointPrefix = "accounts";
		private static string AcsHostUrl = "accesscontrol.windows.net";

		//
		// Hosted add-in configuration
		//
		private static readonly string ClientId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientId")) ? WebConfigurationManager.AppSettings.Get("HostedAppName") : WebConfigurationManager.AppSettings.Get("ClientId");
		private static readonly string IssuerId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("IssuerId")) ? ClientId : WebConfigurationManager.AppSettings.Get("IssuerId");
		private static readonly string HostedAppHostNameOverride = WebConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");
		private static readonly string HostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");
		private static readonly string ClientSecret = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientSecret")) ? WebConfigurationManager.AppSettings.Get("HostedAppSigningKey") : WebConfigurationManager.AppSettings.Get("ClientSecret");
		private static readonly string SecondaryClientSecret = WebConfigurationManager.AppSettings.Get("SecondaryClientSecret");
		private static readonly string Realm = WebConfigurationManager.AppSettings.Get("Realm");
		private static readonly string ServiceNamespace = WebConfigurationManager.AppSettings.Get("Realm");

		private static readonly string ClientSigningCertificatePath = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePath");
		private static readonly string ClientSigningCertificatePassword = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePassword");
		private static readonly X509Certificate2 ClientCertificate = (string.IsNullOrEmpty(ClientSigningCertificatePath) || string.IsNullOrEmpty(ClientSigningCertificatePassword)) ? null : new X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword);
		private static readonly X509SigningCredentials SigningCredentials = (ClientCertificate == null) ? null : new X509SigningCredentials(ClientCertificate, SecurityAlgorithms.RsaSha256Signature, SecurityAlgorithms.Sha256Digest);

		#endregion

		#region メソッド

		/// <summary>
		/// Retrieves the context token string from the specified request by looking for well-known parameter names in the 
		/// POSTed form parameters and the querystring. Returns null if no context token is found.
		/// </summary>
		/// <param name="request">HttpRequest in which to look for a context token</param>
		/// <returns>The context token string</returns>
		public static string GetContextTokenFromRequest(HttpRequest request)
			=> GetContextTokenFromRequest(new HttpRequestWrapper(request));

		/// <summary>
		/// Retrieves the context token string from the specified request by looking for well-known parameter names in the 
		/// POSTed form parameters and the querystring. Returns null if no context token is found.
		/// </summary>
		/// <param name="request">HttpRequest in which to look for a context token</param>
		/// <returns>The context token string</returns>
		public static string GetContextTokenFromRequest(HttpRequestBase request) {
			var paramNames = new[] {
				"AppContext",
				"AppContextToken",
				"AccessToken",
				"SPAppToken"
			};

			foreach (var paramName in paramNames) {
				if (!string.IsNullOrEmpty(request.Form[paramName])) {
					return request.Form[paramName];
				}

				if (!string.IsNullOrEmpty(request.QueryString[paramName])) {
					return request.QueryString[paramName];
				}
			}

			return null;
		}

		/// <summary>
		/// Validate that a specified context token string is intended for this application based on the parameters 
		/// specified in web.config. Parameters used from web.config used for validation include ClientId, 
		/// HostedAppHostNameOverride, HostedAppHostName, ClientSecret, and Realm (if it is specified). If HostedAppHostNameOverride is present,
		/// it will be used for validation. Otherwise, if the <paramref name="appHostName"/> is not 
		/// null, it is used for validation instead of the web.config's HostedAppHostName. If the token is invalid, an 
		/// exception is thrown. If the token is valid, TokenHelper's static STS metadata url is updated based on the token contents
		/// and a JsonWebSecurityToken based on the context token is returned.
		/// </summary>
		/// <param name="contextTokenString">The context token to validate</param>
		/// <param name="appHostName">The URL authority, consisting of  Domain Name System (DNS) host name or IP address and the port number, to use for token audience validation.
		/// If null, HostedAppHostName web.config setting is used instead. HostedAppHostNameOverride web.config setting, if present, will be used 
		/// for validation instead of <paramref name="appHostName"/> .</param>
		/// <returns>A JsonWebSecurityToken based on the context token.</returns>
		public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null) {
			var tokenHandler = CreateJsonWebSecurityTokenHandler();
			var securityToken = tokenHandler.ReadToken(contextTokenString);
			var jsonToken = securityToken as JsonWebSecurityToken;
			var token = SharePointContextToken.Create(jsonToken);

			var stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
			var firstDot = stsAuthority.IndexOf('.');

			GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
			AcsHostUrl = stsAuthority.Substring(firstDot + 1);

			tokenHandler.ValidateToken(jsonToken);

			var acceptableAudiences = GetAcceptableAudiences(appHostName);

			var validationSuccessful = false;
			var realm = Realm ?? token.Realm;
			foreach (var audience in acceptableAudiences) {
				string principal = GetFormattedPrincipal(ClientId, audience, realm);
				if (StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal)) {
					validationSuccessful = true;
					break;
				}
			}

			if (!validationSuccessful) {
				throw new AudienceUriValidationFailedException(
					string.Format(CultureInfo.CurrentCulture,
					"\"{0}\" is not the intended audience \"{1}\"", string.Join(";", acceptableAudiences), token.Audience));
			}

			return token;
		}

		private static string[] GetAcceptableAudiences(string appHostName) {
			if (!string.IsNullOrEmpty(HostedAppHostNameOverride)) {
				return HostedAppHostNameOverride.Split(';');
			} else if (appHostName == null) {
				return new[] { HostedAppHostName };
			} else {
				return new[] { appHostName };
			}
		}

		/// <summary>
		/// Retrieves an access token from ACS to call the source of the specified context token at the specified 
		/// targetHost. The targetHost must be registered for the principal that sent the context token.
		/// </summary>
		/// <param name="contextToken">Context token issued by the intended access token audience</param>
		/// <param name="targetHost">Url authority of the target principal</param>
		/// <returns>An access token with an audience matching the context token's source</returns>
		public static OAuth2AccessTokenResponse GetAccessToken(SharePointContextToken contextToken, string targetHost) {
			var targetPrincipalName = contextToken.TargetPrincipalName;

			// Extract the refreshToken from the context token
			var refreshToken = contextToken.RefreshToken;

			if (string.IsNullOrEmpty(refreshToken)) {
				return null;
			}

			var targetRealm = Realm ?? contextToken.Realm;

			return GetAccessToken(refreshToken, targetPrincipalName, targetHost, targetRealm);
		}

		/// <summary>
		/// Uses the specified authorization code to retrieve an access token from ACS to call the specified principal 
		/// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
		/// null, the "Realm" setting in web.config will be used instead.
		/// </summary>
		/// <param name="authorizationCode">Authorization code to exchange for access token</param>
		/// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
		/// <param name="targetHost">Url authority of the target principal</param>
		/// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
		/// <param name="redirectUri">Redirect URI registered for this add-in</param>
		/// <returns>An access token with an audience of the target principal</returns>
		public static OAuth2AccessTokenResponse GetAccessToken(string authorizationCode, string targetPrincipalName, string targetHost, string targetRealm, Uri redirectUri) {
			if (targetRealm == null) {
				targetRealm = Realm;
			}

			var resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
			var clientId = GetFormattedPrincipal(ClientId, null, targetRealm);

			// Get token
			var oauth2Response = GetOauth2Response(targetRealm, resource, clientId, authorizationCode, redirectUri);

			return oauth2Response;
		}

		/// <summary>
		/// Uses the specified refresh token to retrieve an access token from ACS to call the specified principal 
		/// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
		/// null, the "Realm" setting in web.config will be used instead.
		/// </summary>
		/// <param name="refreshToken">Refresh token to exchange for access token</param>
		/// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
		/// <param name="targetHost">Url authority of the target principal</param>
		/// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
		/// <returns>An access token with an audience of the target principal</returns>
		public static OAuth2AccessTokenResponse GetAccessToken(string refreshToken, string targetPrincipalName, string targetHost, string targetRealm) {
			if (targetRealm == null) {
				targetRealm = Realm;
			}

			var resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
			var clientId = GetFormattedPrincipal(ClientId, null, targetRealm);

			return GetOauth2Response(targetRealm, resource, clientId, refreshToken);
		}

		/// <summary>
		/// Retrieves an app-only access token from ACS to call the specified principal 
		/// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
		/// null, the "Realm" setting in web.config will be used instead.
		/// </summary>
		/// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
		/// <param name="targetHost">Url authority of the target principal</param>
		/// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
		/// <returns>An access token with an audience of the target principal</returns>
		public static OAuth2AccessTokenResponse GetAppOnlyAccessToken(string targetPrincipalName, string targetHost, string targetRealm) {
			if (targetRealm == null) {
				targetRealm = Realm;
			}

			var resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
			var clientId = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm);

			var oauth2Response = GetOauth2Response(targetRealm, resource, clientId);
			return oauth2Response;
		}

		#region GetOauth2Response

		private static OAuth2AccessTokenResponse GetOauth2Response(string targetRealm, string resource, string clientId, string authorizationCode, Uri redirectUri) {
			return GetOauth2Response(targetRealm, c => OAuth2MessageFactory.CreateAccessTokenRequestWithAuthorizationCode(clientId, c, authorizationCode, redirectUri, resource));
		}

		private static OAuth2AccessTokenResponse GetOauth2Response(string targetRealm, string resource, string clientId, string refreshToken) {
			return GetOauth2Response(targetRealm, c => OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(clientId, c, refreshToken, resource));
		}

		private static OAuth2AccessTokenResponse GetOauth2Response(string targetRealm, string resource, string clientId) {
			return GetOauth2Response(targetRealm, c => {
				var request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, c, resource);
				request.Resource = resource;
				return request;
			});
		}

		private static OAuth2AccessTokenResponse GetOauth2Response(string targetRealm, Func<string, OAuth2AccessTokenRequest> createRequest) {
			var client = new OAuth2S2SClient();
			try {
				var oauth2Request = createRequest?.Invoke(ClientSecret);
				return client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
			} catch (RequestFailedException ex) {
				if (!string.IsNullOrEmpty(SecondaryClientSecret)) {
					var oauth2Request = createRequest?.Invoke(SecondaryClientSecret);
					return client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
				} else {
					throw ex;
				}
			} catch (WebException ex) {
				using (var sr = new StreamReader(ex.Response.GetResponseStream())) {
					var responseText = sr.ReadToEnd();
					throw new WebException($"{ex.Message} - {responseText}", ex);
				}
			}
		}

		#endregion

		#region CreateRemoteEventReceiverClientContext

		/// <summary>
		/// Creates a client context based on the properties of a remote event receiver
		/// </summary>
		/// <param name="properties">Properties of a remote event receiver</param>
		/// <returns>A ClientContext ready to call the web where the event originated</returns>
		public static ClientContext CreateRemoteEventReceiverClientContext(SPRemoteEventProperties properties) {
			var uri = properties.GetSharepointUri();
			if (uri == null) {
				return null;
			}

			return uri.GetClientContext(properties);
		}

		private static Uri GetSharepointUri(this SPRemoteEventProperties @this) {
			if (@this?.ListEventProperties != null) {
				return new Uri(@this.ListEventProperties.WebUrl);
			} else if (@this?.ItemEventProperties != null) {
				return new Uri(@this.ItemEventProperties.WebUrl);
			} else if (@this?.WebEventProperties != null) {
				return new Uri(@this.WebEventProperties.FullUrl);
			}

			return null;
		}

		#endregion

		/// <summary>
		/// Creates a client context based on the properties of an add-in event
		/// </summary>
		/// <param name="properties">Properties of an add-in event</param>
		/// <param name="useAppWeb">True to target the app web, false to target the host web</param>
		/// <returns>A ClientContext ready to call the app web or the parent web</returns>
		public static ClientContext CreateAppEventClientContext(SPRemoteEventProperties properties, bool useAppWeb) {
			var uri = useAppWeb
				? properties?.AppEventProperties?.AppWebFullUrl
				: properties?.AppEventProperties?.HostWebFullUrl;

			if (uri == null) {
				return null;
			}

			return uri.GetClientContext(properties);
		}

		private static ClientContext GetClientContext(this Uri uri, SPRemoteEventProperties properties) {
			return IsHighTrustApp
				? GetS2SClientContextWithWindowsIdentity(uri, null)
				: CreateAcsClientContextForUrl(properties, uri);
		}

		/// <summary>
		/// Retrieves an access token from ACS using the specified authorization code, and uses that access token to 
		/// create a client context
		/// </summary>
		/// <param name="targetUrl">Url of the target SharePoint site</param>
		/// <param name="authorizationCode">Authorization code to use when retrieving the access token from ACS</param>
		/// <param name="redirectUri">Redirect URI registered for this add-in</param>
		/// <returns>A ClientContext ready to call targetUrl with a valid access token</returns>
		public static ClientContext GetClientContextWithAuthorizationCode(
			string targetUrl,
			string authorizationCode,
			Uri redirectUri) {
			return GetClientContextWithAuthorizationCode(targetUrl, SharePointPrincipal, authorizationCode, GetRealmFromTargetUrl(new Uri(targetUrl)), redirectUri);
		}

		/// <summary>
		/// Retrieves an access token from ACS using the specified authorization code, and uses that access token to 
		/// create a client context
		/// </summary>
		/// <param name="targetUrl">Url of the target SharePoint site</param>
		/// <param name="targetPrincipalName">Name of the target SharePoint principal</param>
		/// <param name="authorizationCode">Authorization code to use when retrieving the access token from ACS</param>
		/// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
		/// <param name="redirectUri">Redirect URI registered for this add-in</param>
		/// <returns>A ClientContext ready to call targetUrl with a valid access token</returns>
		public static ClientContext GetClientContextWithAuthorizationCode(
			string targetUrl,
			string targetPrincipalName,
			string authorizationCode,
			string targetRealm,
			Uri redirectUri) {
			Uri targetUri = new Uri(targetUrl);

			string accessToken =
				GetAccessToken(authorizationCode, targetPrincipalName, targetUri.Authority, targetRealm, redirectUri).AccessToken;

			return GetClientContextWithAccessToken(targetUrl, accessToken);
		}

		/// <summary>
		/// Uses the specified access token to create a client context
		/// </summary>
		/// <param name="targetUrl">Url of the target SharePoint site</param>
		/// <param name="accessToken">Access token to be used when calling the specified targetUrl</param>
		/// <returns>A ClientContext ready to call targetUrl with the specified access token</returns>
		public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken) {
			ClientContext clientContext = new ClientContext(targetUrl);

			clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
			clientContext.FormDigestHandlingEnabled = false;
			clientContext.ExecutingWebRequest +=
				delegate (object oSender, WebRequestEventArgs webRequestEventArgs) {
					webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
						"Bearer " + accessToken;
				};

			return clientContext;
		}

		/// <summary>
		/// Retrieves an access token from ACS using the specified context token, and uses that access token to create
		/// a client context
		/// </summary>
		/// <param name="targetUrl">Url of the target SharePoint site</param>
		/// <param name="contextTokenString">Context token received from the target SharePoint site</param>
		/// <param name="appHostUrl">Url authority of the hosted add-in.  If this is null, the value in the HostedAppHostName
		/// of web.config will be used instead</param>
		/// <returns>A ClientContext ready to call targetUrl with a valid access token</returns>
		public static ClientContext GetClientContextWithContextToken(string targetUrl, string contextTokenString, string appHostUrl) {
			var contextToken = ReadAndValidateContextToken(contextTokenString, appHostUrl);
			var targetUri = new Uri(targetUrl);
			var accessToken = GetAccessToken(contextToken, targetUri.Authority).AccessToken;

			return GetClientContextWithAccessToken(targetUrl, accessToken);
		}

		/// <summary>
		/// Returns the SharePoint url to which the add-in should redirect the browser to request consent and get back
		/// an authorization code.
		/// </summary>
		/// <param name="contextUrl">Absolute Url of the SharePoint site</param>
		/// <param name="scope">Space-delimited permissions to request from the SharePoint site in "shorthand" format 
		/// (e.g. "Web.Read Site.Write")</param>
		/// <returns>Url of the SharePoint site's OAuth authorization page</returns>
		public static string GetAuthorizationUrl(string contextUrl, string scope) {
			return string.Format(
				"{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code",
				EnsureTrailingSlash(contextUrl),
				AuthorizationPage,
				ClientId,
				scope);
		}

		/// <summary>
		/// Returns the SharePoint url to which the add-in should redirect the browser to request consent and get back
		/// an authorization code.
		/// </summary>
		/// <param name="contextUrl">Absolute Url of the SharePoint site</param>
		/// <param name="scope">Space-delimited permissions to request from the SharePoint site in "shorthand" format
		/// (e.g. "Web.Read Site.Write")</param>
		/// <param name="redirectUri">Uri to which SharePoint should redirect the browser to after consent is 
		/// granted</param>
		/// <returns>Url of the SharePoint site's OAuth authorization page</returns>
		public static string GetAuthorizationUrl(string contextUrl, string scope, string redirectUri) {
			return string.Format(
				"{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code&redirect_uri={4}",
				EnsureTrailingSlash(contextUrl),
				AuthorizationPage,
				ClientId,
				scope,
				redirectUri);
		}

		/// <summary>
		/// Returns the SharePoint url to which the add-in should redirect the browser to request a new context token.
		/// </summary>
		/// <param name="contextUrl">Absolute Url of the SharePoint site</param>
		/// <param name="redirectUri">Uri to which SharePoint should redirect the browser to with a context token</param>
		/// <returns>Url of the SharePoint site's context token redirect page</returns>
		public static string GetAppContextTokenRequestUrl(string contextUrl, string redirectUri) {
			return string.Format(
				"{0}{1}?client_id={2}&redirect_uri={3}",
				EnsureTrailingSlash(contextUrl),
				RedirectPage,
				ClientId,
				redirectUri);
		}

		/// <summary>
		/// Retrieves an S2S access token signed by the application's private certificate on behalf of the specified 
		/// WindowsIdentity and intended for the SharePoint at the targetApplicationUri. If no Realm is specified in 
		/// web.config, an auth challenge will be issued to the targetApplicationUri to discover it.
		/// </summary>
		/// <param name="targetApplicationUri">Url of the target SharePoint site</param>
		/// <param name="identity">Windows identity of the user on whose behalf to create the access token</param>
		/// <returns>An access token with an audience of the target principal</returns>
		public static string GetS2SAccessTokenWithWindowsIdentity(Uri targetApplicationUri, WindowsIdentity identity) {
			var realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;
			var claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;
			return GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
		}

		/// <summary>
		/// Retrieves an S2S client context with an access token signed by the application's private certificate on 
		/// behalf of the specified WindowsIdentity and intended for application at the targetApplicationUri using the 
		/// targetRealm. If no Realm is specified in web.config, an auth challenge will be issued to the 
		/// targetApplicationUri to discover it.
		/// </summary>
		/// <param name="targetApplicationUri">Url of the target SharePoint site</param>
		/// <param name="identity">Windows identity of the user on whose behalf to create the access token</param>
		/// <returns>A ClientContext using an access token with an audience of the target application</returns>
		public static ClientContext GetS2SClientContextWithWindowsIdentity(Uri targetApplicationUri, WindowsIdentity identity) {
			var realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;
			var claims = (identity != null) ? GetClaimsWithWindowsIdentity(identity) : null;
			var accessToken = GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);

			return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
		}

		/// <summary>
		/// Get authentication realm from SharePoint
		/// </summary>
		/// <param name="targetApplicationUri">Url of the target SharePoint site</param>
		/// <returns>String representation of the realm GUID</returns>
		public static string GetRealmFromTargetUrl(Uri targetApplicationUri) {
			var request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
			request.Headers.Add("Authorization: Bearer ");

			try {
				using (request.GetResponse()) {
				}
			} catch (WebException e) {
				if (e.Response == null) {
					return null;
				}

				var bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
				if (string.IsNullOrEmpty(bearerResponseHeader)) {
					return null;
				}

				const string bearer = "Bearer realm=\"";
				int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
				if (bearerIndex < 0) {
					return null;
				}

				int realmIndex = bearerIndex + bearer.Length;

				if (bearerResponseHeader.Length >= realmIndex + 36) {
					var targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

					Guid realmGuid;
					if (Guid.TryParse(targetRealm, out realmGuid)) {
						return targetRealm;
					}
				}
			}
			return null;
		}

		/// <summary>
		/// Determines if this is a high trust add-in.
		/// </summary>
		/// <returns>True if this is a high trust add-in.</returns>
		public static bool IsHighTrustApp => SigningCredentials != null;

		/// <summary>
		/// Ensures that the specified URL ends with '/' if it is not null or empty.
		/// </summary>
		/// <param name="url">The url.</param>
		/// <returns>The url ending with '/' if it is not null or empty.</returns>
		public static string EnsureTrailingSlash(string url) {
			if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/') {
				return url + "/";
			}

			return url;
		}

		#endregion

		#region private methods

		private static ClientContext CreateAcsClientContextForUrl(SPRemoteEventProperties properties, Uri sharepointUrl) {
			string contextTokenString = properties.ContextToken;

			if (string.IsNullOrEmpty(contextTokenString)) {
				return null;
			}

			SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host);
			string accessToken = GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

			return GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
		}

		public static string GetAcsMetadataEndpointUrl() {
			return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
		}

		private static string GetFormattedPrincipal(string principalName, string hostName, string realm) {
			if (!string.IsNullOrEmpty(hostName)) {
				return string.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
			}

			return string.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
		}

		private static string GetAcsPrincipalName(string realm) {
			return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
		}

		private static string GetAcsGlobalEndpointUrl() {
			return string.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
		}

		private static JsonWebSecurityTokenHandler CreateJsonWebSecurityTokenHandler() {
			JsonWebSecurityTokenHandler handler = new JsonWebSecurityTokenHandler();
			handler.Configuration = new SecurityTokenHandlerConfiguration();
			handler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Never);
			handler.Configuration.CertificateValidator = X509CertificateValidator.None;

			List<byte[]> securityKeys = new List<byte[]>();
			securityKeys.Add(Convert.FromBase64String(ClientSecret));
			if (!string.IsNullOrEmpty(SecondaryClientSecret)) {
				securityKeys.Add(Convert.FromBase64String(SecondaryClientSecret));
			}

			List<SecurityToken> securityTokens = new List<SecurityToken>();
			securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

			handler.Configuration.IssuerTokenResolver =
				SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
				new ReadOnlyCollection<SecurityToken>(securityTokens),
				false);
			SymmetricKeyIssuerNameRegistry issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
			foreach (byte[] securitykey in securityKeys) {
				issuerNameRegistry.AddTrustedIssuer(securitykey, GetAcsPrincipalName(ServiceNamespace));
			}
			handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
			return handler;
		}

		private static string GetS2SAccessTokenWithClaims(
			string targetApplicationHostName,
			string targetRealm,
			IEnumerable<JsonWebTokenClaim> claims) {
			return IssueToken(
				ClientId,
				IssuerId,
				targetRealm,
				SharePointPrincipal,
				targetRealm,
				targetApplicationHostName,
				true,
				claims,
				claims == null);
		}

		private static JsonWebTokenClaim[] GetClaimsWithWindowsIdentity(WindowsIdentity identity) {
			JsonWebTokenClaim[] claims = new JsonWebTokenClaim[]
			{
				new JsonWebTokenClaim(NameIdentifierClaimType, identity.User.Value.ToLower()),
				new JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")
			};
			return claims;
		}

		private static string IssueToken(
			string sourceApplication,
			string issuerApplication,
			string sourceRealm,
			string targetApplication,
			string targetRealm,
			string targetApplicationHostName,
			bool trustedForDelegation,
			IEnumerable<JsonWebTokenClaim> claims,
			bool appOnly = false) {
			if (null == SigningCredentials) {
				throw new InvalidOperationException("SigningCredentials was not initialized");
			}

			#region Actor token

			string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
			string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
			string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

			List<JsonWebTokenClaim> actorClaims = new List<JsonWebTokenClaim>();
			actorClaims.Add(new JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, nameid));
			if (trustedForDelegation && !appOnly) {
				actorClaims.Add(new JsonWebTokenClaim(TrustedForImpersonationClaimType, "true"));
			}

			// Create token
			JsonWebSecurityToken actorToken = new JsonWebSecurityToken(
				issuer: issuer,
				audience: audience,
				validFrom: DateTime.UtcNow,
				validTo: DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
				signingCredentials: SigningCredentials,
				claims: actorClaims);

			string actorTokenString = new JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken);

			if (appOnly) {
				// App-only token is the same as actor token for delegated case
				return actorTokenString;
			}

			#endregion Actor token

			#region Outer token

			List<JsonWebTokenClaim> outerClaims = null == claims ? new List<JsonWebTokenClaim>() : new List<JsonWebTokenClaim>(claims);
			outerClaims.Add(new JsonWebTokenClaim(ActorTokenClaimType, actorTokenString));

			JsonWebSecurityToken jsonToken = new JsonWebSecurityToken(
				nameid, // outer token issuer should match actor token nameid
				audience,
				DateTime.UtcNow,
				DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
				outerClaims);

			string accessToken = new JsonWebSecurityTokenHandler().WriteTokenAsString(jsonToken);

			#endregion Outer token

			return accessToken;
		}

		#endregion

		#region AcsMetadataParser

		#endregion
	}
}
