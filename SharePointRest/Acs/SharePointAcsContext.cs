using System;
using System.Net;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Models;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Acs {
	/// <summary>
	/// Encapsulates all the information from SharePoint in ACS mode.
	/// </summary>
	public class SharePointAcsContext : SharePointContext {
		#region フィールド

		private readonly string _contextToken;
		private readonly SharePointContextToken _contextTokenObj;

		#endregion

		#region コンストラクタ

		public SharePointAcsContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, string contextToken, SharePointContextToken contextTokenObj)
			: base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber) {
			if (string.IsNullOrEmpty(contextToken)) {
				throw new ArgumentNullException("contextToken");
			}

			if (contextTokenObj == null) {
				throw new ArgumentNullException("contextTokenObj");
			}

			this._contextToken = contextToken;
			this._contextTokenObj = contextTokenObj;
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// The context token.
		/// </summary>
		public string ContextToken
			=> this._contextTokenObj.ValidTo > DateTime.UtcNow ? this._contextToken : null;

		/// <summary>
		/// The context token's "CacheKey" claim.
		/// </summary>
		public string CacheKey
			=> this._contextTokenObj.ValidTo > DateTime.UtcNow ? this._contextTokenObj.CacheKey : null;

		/// <summary>
		/// The context token's "refreshtoken" claim.
		/// </summary>
		public string RefreshToken
			=> this._contextTokenObj.ValidTo > DateTime.UtcNow ? this._contextTokenObj.RefreshToken : null;

		public override string UserAccessTokenForSPHost
			=> GetAccessTokenString(ref this.userAccessTokenForSPHost, () => TokenHelper.GetAccessToken(this._contextTokenObj, this.SPHostUrl.Authority));

		public override string UserAccessTokenForSPAppWeb
			=> (this.SPAppWebUrl == null)
				? null
				: GetAccessTokenString(ref this.userAccessTokenForSPAppWeb, () => TokenHelper.GetAccessToken(this._contextTokenObj, this.SPAppWebUrl.Authority));

		public override string AppOnlyAccessTokenForSPHost
			=> GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost, () => TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, this.SPHostUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPHostUrl)));

		public override string AppOnlyAccessTokenForSPAppWeb
			=> (this.SPAppWebUrl == null)
				? null
				: GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb, () => TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, this.SPAppWebUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPAppWebUrl)));

		#endregion

		#region メソッド

		/// <summary>
		/// Ensures the access token is valid and returns it.
		/// </summary>
		/// <param name="accessToken">The access token to verify.</param>
		/// <param name="tokenRenewalHandler">The token renewal handler.</param>
		/// <returns>The access token string.</returns>
		private static string GetAccessTokenString(ref AccessToken accessToken, Func<OAuth2AccessTokenResponse> tokenRenewalHandler) {
			RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

			return accessToken.IsValid ? accessToken.Item1 : null;
		}

		/// <summary>
		/// Renews the access token if it is not valid.
		/// </summary>
		/// <param name="accessToken">The access token to renew.</param>
		/// <param name="tokenRenewalHandler">The token renewal handler.</param>
		private static void RenewAccessTokenIfNeeded(ref AccessToken accessToken, Func<OAuth2AccessTokenResponse> tokenRenewalHandler) {
			if (accessToken.IsValid) {
				return;
			}

			try {
				var oAuth2AccessTokenResponse = tokenRenewalHandler?.Invoke();

				var expiresOn = oAuth2AccessTokenResponse.ExpiresOn;

				if ((expiresOn - oAuth2AccessTokenResponse.NotBefore) > AccessTokenLifetimeTolerance) {
					// Make the access token get renewed a bit earlier than the time when it expires
					// so that the calls to SharePoint with it will have enough time to complete successfully.
					expiresOn -= AccessTokenLifetimeTolerance;
				}

				accessToken = new AccessToken(oAuth2AccessTokenResponse.AccessToken, expiresOn);
			} catch (WebException) {
			}
		}

		#endregion
	}
}