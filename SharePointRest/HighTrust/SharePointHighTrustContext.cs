using System;
using System.Security.Principal;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Models;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.HighTrust {
	/// <summary>
	/// Encapsulates all the information from SharePoint in HighTrust mode.
	/// </summary>
	public class SharePointHighTrustContext : SharePointContext {
		#region フィールド

		private readonly WindowsIdentity logonUserIdentity;

		#endregion

		#region コンストラクタ

		public SharePointHighTrustContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, WindowsIdentity logonUserIdentity)
			: base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber) {
			if (logonUserIdentity == null) {
				throw new ArgumentNullException("logonUserIdentity");
			}

			this.logonUserIdentity = logonUserIdentity;
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// The Windows identity for the current user.
		/// </summary>
		public WindowsIdentity LogonUserIdentity
			=> this.logonUserIdentity;

		public override string UserAccessTokenForSPHost
			=> GetAccessTokenString(ref this.userAccessTokenForSPHost, () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, this.LogonUserIdentity));

		public override string UserAccessTokenForSPAppWeb
			=> (this.SPAppWebUrl == null)
				? null
				: GetAccessTokenString(ref this.userAccessTokenForSPAppWeb, () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, this.LogonUserIdentity));

		public override string AppOnlyAccessTokenForSPHost
			=> GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost, () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, null));

		public override string AppOnlyAccessTokenForSPAppWeb
			=> (this.SPAppWebUrl == null)
				? null
				: GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb, () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, null));

		#endregion

		#region メソッド

		/// <summary>
		/// Ensures the access token is valid and returns it.
		/// </summary>
		/// <param name="accessToken">The access token to verify.</param>
		/// <param name="tokenRenewalHandler">The token renewal handler.</param>
		/// <returns>The access token string.</returns>
		private static string GetAccessTokenString(ref AccessToken accessToken, Func<string> tokenRenewalHandler) {
			RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

			return accessToken.IsValid ? accessToken.Item1 : null;
		}

		/// <summary>
		/// Renews the access token if it is not valid.
		/// </summary>
		/// <param name="accessToken">The access token to renew.</param>
		/// <param name="tokenRenewalHandler">The token renewal handler.</param>
		private static void RenewAccessTokenIfNeeded(ref AccessToken accessToken, Func<string> tokenRenewalHandler) {
			if (accessToken.IsValid) {
				return;
			}

			var expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);

			if (TokenHelper.HighTrustAccessTokenLifetime > AccessTokenLifetimeTolerance) {
				// Make the access token get renewed a bit earlier than the time when it expires
				// so that the calls to SharePoint with it will have enough time to complete successfully.
				expiresOn -= AccessTokenLifetimeTolerance;
			}

			accessToken = new AccessToken(tokenRenewalHandler(), expiresOn);
		}

		#endregion
	}
}