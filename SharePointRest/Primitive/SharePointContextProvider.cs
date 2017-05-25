using System;
using System.Web;
using Microsoft.IdentityModel.Tokens;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Acs;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Enums;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Extensions;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.HighTrust;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;
using static SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Const.ConstString;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive {
	/// <summary>
	/// Provides SharePointContext instances.
	/// </summary>
	public abstract class SharePointContextProvider {
		#region コンストラクタ

		/// <summary>
		/// Initializes the default SharePointContextProvider instance.
		/// </summary>
		static SharePointContextProvider() {
			var isHighTrustApp = TokenHelper.IsHighTrustApp;
			if (!isHighTrustApp) {
				Current = new SharePointAcsContextProvider();
			} else {
				Current = new SharePointHighTrustContextProvider();
			}
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// 現在の SharePointContextProvider インスタンスを取得します。
		/// </summary>
		public static SharePointContextProvider Current { get; protected set; }

		#endregion

		#region メソッド

		/// <summary>
		/// Registers the specified SharePointContextProvider instance as current.
		/// It should be called by Application_Start() in Global.asax.
		/// </summary>
		/// <param name="provider">The SharePointContextProvider to be set as current.</param>
		public static void Register(SharePointContextProvider provider) {
			if (provider == null) {
				throw new ArgumentNullException(nameof(provider));
			}

			Current = provider;
		}

		/// <summary>
		/// Checks if it is necessary to redirect to SharePoint for user to authenticate.
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="redirectUrl">The redirect url to SharePoint if the status is ShouldRedirect. <c>Null</c> if the status is Ok or CanNotRedirect.</param>
		/// <returns>Redirection status.</returns>
		public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, out Uri redirectUrl) {
			if (httpContext == null) {
				throw new ArgumentNullException(nameof(httpContext));
			}

			redirectUrl = null;
			var contextTokenExpired = false;

			try {
				if (Current.GetSharePointContext(httpContext) != null) {
					return RedirectionStatus.Ok;
				}
			} catch (SecurityTokenExpiredException) {
				contextTokenExpired = true;
			}

			if (!string.IsNullOrEmpty(httpContext.Request.QueryString[SPHasRedirectedToSharePointKey]) && !contextTokenExpired) {
				return RedirectionStatus.CanNotRedirect;
			}

			var spHostUrl = httpContext.Request.GetSPHostUrl();

			if (spHostUrl == null) {
				return RedirectionStatus.CanNotRedirect;
			}

			if (StringComparer.OrdinalIgnoreCase.Equals(httpContext.Request.HttpMethod, "POST")) {
				return RedirectionStatus.CanNotRedirect;
			}

			var requestUrl = httpContext.Request.Url;

			var queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query);

			// Removes the values that are included in {StandardTokens}, as {StandardTokens} will be inserted at the beginning of the query string.
			queryNameValueCollection.Remove(SPHostUrlKey);
			queryNameValueCollection.Remove(SPAppWebUrlKey);
			queryNameValueCollection.Remove(SPLanguageKey);
			queryNameValueCollection.Remove(SPClientTagKey);
			queryNameValueCollection.Remove(SPProductNumberKey);

			// Adds SPHasRedirectedToSharePoint=1.
			queryNameValueCollection.Add(SPHasRedirectedToSharePointKey, "1");

			var returnUrlBuilder = new UriBuilder(requestUrl);
			returnUrlBuilder.Query = queryNameValueCollection.ToString();

			// Inserts StandardTokens.
			var returnUrlString = returnUrlBuilder.Uri.AbsoluteUri;
			returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?") + 1, $"{{{StandardTokens}}}&");

			// Constructs redirect url.
			var redirectUrlString = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.AbsoluteUri, Uri.EscapeDataString(returnUrlString));

			redirectUrl = new Uri(redirectUrlString, UriKind.Absolute);

			return RedirectionStatus.ShouldRedirect;
		}

		/// <summary>
		/// Checks if it is necessary to redirect to SharePoint for user to authenticate.
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="redirectUrl">The redirect url to SharePoint if the status is ShouldRedirect. <c>Null</c> if the status is Ok or CanNotRedirect.</param>
		/// <returns>Redirection status.</returns>
		public static RedirectionStatus CheckRedirectionStatus(HttpContext httpContext, out Uri redirectUrl) {
			return CheckRedirectionStatus(new HttpContextWrapper(httpContext), out redirectUrl);
		}

		/// <summary>
		/// Creates a SharePointContext instance with the specified HTTP request.
		/// </summary>
		/// <param name="httpRequest">The HTTP request.</param>
		/// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
		public SharePointContext CreateSharePointContext(HttpRequestBase httpRequest) {
			if (httpRequest == null) {
				throw new ArgumentNullException("httpRequest");
			}

			// SPHostUrl
			Uri spHostUrl = httpRequest.GetSPHostUrl();
			if (spHostUrl == null) {
				return null;
			}

			// SPAppWebUrl
			string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SPAppWebUrlKey]);
			Uri spAppWebUrl;
			if (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl) ||
				!(spAppWebUrl.Scheme == Uri.UriSchemeHttp || spAppWebUrl.Scheme == Uri.UriSchemeHttps)) {
				spAppWebUrl = null;
			}

			// SPLanguage
			string spLanguage = httpRequest.QueryString[SPLanguageKey];
			if (string.IsNullOrEmpty(spLanguage)) {
				return null;
			}

			// SPClientTag
			string spClientTag = httpRequest.QueryString[SPClientTagKey];
			if (string.IsNullOrEmpty(spClientTag)) {
				return null;
			}

			// SPProductNumber
			string spProductNumber = httpRequest.QueryString[SPProductNumberKey];
			if (string.IsNullOrEmpty(spProductNumber)) {
				return null;
			}

			return CreateSharePointContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, httpRequest);
		}

		/// <summary>
		/// Creates a SharePointContext instance with the specified HTTP request.
		/// </summary>
		/// <param name="httpRequest">The HTTP request.</param>
		/// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
		public SharePointContext CreateSharePointContext(HttpRequest httpRequest) {
			return CreateSharePointContext(new HttpRequestWrapper(httpRequest));
		}

		/// <summary>
		/// Gets a SharePointContext instance associated with the specified HTTP context.
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <returns>The SharePointContext instance. Returns <c>null</c> if not found and a new instance can't be created.</returns>
		public SharePointContext GetSharePointContext(HttpContextBase httpContext) {
			if (httpContext == null) {
				throw new ArgumentNullException(nameof(httpContext));
			}

			var spHostUrl = httpContext.Request.GetSPHostUrl();
			if (spHostUrl == null) {
				return null;
			}

			var spContext = LoadSharePointContext(httpContext);

			if (spContext == null || !ValidateSharePointContext(httpContext, spContext)) {
				spContext = CreateSharePointContext(httpContext.Request);

				if (spContext != null) {
					SaveSharePointContext(httpContext, spContext);
				}
			}

			return spContext;
		}

		/// <summary>
		/// Gets a SharePointContext instance associated with the specified HTTP context.
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <returns>The SharePointContext instance. Returns <c>null</c> if not found and a new instance can't be created.</returns>
		public SharePointContext GetSharePointContext(HttpContext httpContext) {
			return GetSharePointContext(new HttpContextWrapper(httpContext));
		}

		#region 抽象メソッド

		/// <summary>
		/// SharePointContext インスタンスを生成します。
		/// </summary>
		/// <param name="spHostUrl">SharePoint ホストの URL</param>
		/// <param name="spAppWebUrl">SharePoint アプリケーション Web URL</param>
		/// <param name="spLanguage">SharePoint の言語</param>
		/// <param name="spClientTag">SharePoint クライアントタグ。</param>
		/// <param name="spProductNumber">SharePoint 製品番号</param>
		/// <param name="httpRequest">HTTP リクエスト</param>
		/// <returns>生成した SharePointContext インスタンスを返します。
		/// エラーが発生した場合は <c>null</c> を返します。</returns>
		protected abstract SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest);

		/// <summary>
		/// 指定された SharePoint コンテキストを指定された HTTP コンテキストで使用できるかどうかを検証します。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="spContext">SharePointContext</param>
		/// <returns>使用できる場合は true を返します。</returns>
		protected abstract bool ValidateSharePointContext(HttpContextBase httpContext, SharePointContext spContext);

		/// <summary>
		/// 指定された HTTP コンテキストに関連付けられた SharePointContext インスタンスを読み込みます。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <returns>SharePointContext インスタンスを返します。
		/// 見つからなければ <c>null</c> を返します。</returns>
		protected abstract SharePointContext LoadSharePointContext(HttpContextBase httpContext);

		/// <summary>
		/// 指定された HTTP コンテキストに関連付けられている指定された SharePointContext インスタンスを保存します。
		/// <c>null</c>は、HTTP コンテキストに関連付けられた SharePointContext インスタンスを消去するために受け入れられます。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="spContext">保存する SharePointContext インスタンス</param>
		protected abstract void SaveSharePointContext(HttpContextBase httpContext, SharePointContext spContext);

		#endregion

		#endregion
	}
}