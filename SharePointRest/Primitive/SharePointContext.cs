using System;
using Microsoft.SharePoint.Client;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Models;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive {
	/// <summary>
	/// SharePoint からのすべての情報をカプセル化します。
	/// </summary>
	public abstract class SharePointContext {
		#region フィールド

		protected static readonly TimeSpan AccessTokenLifetimeTolerance = TimeSpan.FromMinutes(5.0);

		private readonly Uri spHostUrl;
		private readonly Uri spAppWebUrl;
		private readonly string spLanguage;
		private readonly string spClientTag;
		private readonly string spProductNumber;

		// <AccessTokenString, UtcExpiresOn>
		protected AccessToken userAccessTokenForSPHost;
		protected AccessToken userAccessTokenForSPAppWeb;
		protected AccessToken appOnlyAccessTokenForSPHost;
		protected AccessToken appOnlyAccessTokenForSPAppWeb;

		#endregion

		#region プロパティ

		/// <summary>
		/// The SharePoint host url.
		/// </summary>
		public Uri SPHostUrl {
			get { return this.spHostUrl; }
		}

		/// <summary>
		/// The SharePoint app web url.
		/// </summary>
		public Uri SPAppWebUrl {
			get { return this.spAppWebUrl; }
		}

		/// <summary>
		/// The SharePoint language.
		/// </summary>
		public string SPLanguage {
			get { return this.spLanguage; }
		}

		/// <summary>
		/// The SharePoint client tag.
		/// </summary>
		public string SPClientTag {
			get { return this.spClientTag; }
		}

		/// <summary>
		/// The SharePoint product number.
		/// </summary>
		public string SPProductNumber {
			get { return this.spProductNumber; }
		}

		/// <summary>
		/// The user access token for the SharePoint host.
		/// </summary>
		public abstract string UserAccessTokenForSPHost {
			get;
		}

		/// <summary>
		/// The user access token for the SharePoint app web.
		/// </summary>
		public abstract string UserAccessTokenForSPAppWeb {
			get;
		}

		/// <summary>
		/// The app only access token for the SharePoint host.
		/// </summary>
		public abstract string AppOnlyAccessTokenForSPHost {
			get;
		}

		/// <summary>
		/// The app only access token for the SharePoint app web.
		/// </summary>
		public abstract string AppOnlyAccessTokenForSPAppWeb {
			get;
		}

		#endregion

		#region コンストラクタ

		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="spHostUrl">The SharePoint host url.</param>
		/// <param name="spAppWebUrl">The SharePoint app web url.</param>
		/// <param name="spLanguage">The SharePoint language.</param>
		/// <param name="spClientTag">The SharePoint client tag.</param>
		/// <param name="spProductNumber">The SharePoint product number.</param>
		protected SharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber) {
			if (spHostUrl == null) {
				throw new ArgumentNullException("spHostUrl");
			}

			if (string.IsNullOrEmpty(spLanguage)) {
				throw new ArgumentNullException("spLanguage");
			}

			if (string.IsNullOrEmpty(spClientTag)) {
				throw new ArgumentNullException("spClientTag");
			}

			if (string.IsNullOrEmpty(spProductNumber)) {
				throw new ArgumentNullException("spProductNumber");
			}

			this.spHostUrl = spHostUrl;
			this.spAppWebUrl = spAppWebUrl;
			this.spLanguage = spLanguage;
			this.spClientTag = spClientTag;
			this.spProductNumber = spProductNumber;
		}

		#endregion

		#region メソッド

		#region Create

		/// <summary>
		/// Creates a user ClientContext for the SharePoint host.
		/// </summary>
		/// <returns>A ClientContext instance.</returns>
		public ClientContext CreateUserClientContextForSPHost() {
			return CreateClientContext(this.SPHostUrl, this.UserAccessTokenForSPHost);
		}

		/// <summary>
		/// Creates a user ClientContext for the SharePoint app web.
		/// </summary>
		/// <returns>A ClientContext instance.</returns>
		public ClientContext CreateUserClientContextForSPAppWeb() {
			return CreateClientContext(this.SPAppWebUrl, this.UserAccessTokenForSPAppWeb);
		}

		/// <summary>
		/// Creates app only ClientContext for the SharePoint host.
		/// </summary>
		/// <returns>A ClientContext instance.</returns>
		public ClientContext CreateAppOnlyClientContextForSPHost() {
			return CreateClientContext(this.SPHostUrl, this.AppOnlyAccessTokenForSPHost);
		}

		/// <summary>
		/// Creates an app only ClientContext for the SharePoint app web.
		/// </summary>
		/// <returns>A ClientContext instance.</returns>
		public ClientContext CreateAppOnlyClientContextForSPAppWeb() {
			return CreateClientContext(this.SPAppWebUrl, this.AppOnlyAccessTokenForSPAppWeb);
		}

		#endregion

		/// <summary>
		/// Gets the database connection string from SharePoint for autohosted add-in.
		/// This method is deprecated because the autohosted option is no longer available.
		/// </summary>
		[ObsoleteAttribute("This method is deprecated because the autohosted option is no longer available.", true)]
		public string GetDatabaseConnectionString() {
			throw new NotSupportedException("This method is deprecated because the autohosted option is no longer available.");
		}

		/// <summary>
		/// Creates a ClientContext with the specified SharePoint site url and the access token.
		/// </summary>
		/// <param name="spSiteUrl">The site url.</param>
		/// <param name="accessToken">The access token.</param>
		/// <returns>A ClientContext instance.</returns>
		private static ClientContext CreateClientContext(Uri spSiteUrl, string accessToken) {
			if (spSiteUrl != null && !string.IsNullOrEmpty(accessToken)) {
				return TokenHelper.GetClientContextWithAccessToken(spSiteUrl.AbsoluteUri, accessToken);
			}

			return null;
		}

		#endregion
	}

	public static partial class AccessTokenExtension {
	}
}
