using System;
using System.Net;
using System.Web;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Extensions;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;
using static SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Const.ConstString;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Acs {
	/// <summary>
	/// Default provider for SharePointAcsContext.
	/// </summary>
	public class SharePointAcsContextProvider : SharePointContextProvider {
		#region メソッド

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
		protected override SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest) {
			string contextTokenString = TokenHelper.GetContextTokenFromRequest(httpRequest);
			if (string.IsNullOrEmpty(contextTokenString)) {
				return null;
			}

			SharePointContextToken contextToken = null;
			try {
				contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpRequest.Url.Authority);
			} catch (WebException) {
				return null;
			} catch (System.IdentityModel.Tokens.AudienceUriValidationFailedException) {
				return null;
			}

			return new SharePointAcsContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, contextTokenString, contextToken);
		}

		/// <summary>
		/// 指定された SharePoint コンテキストを指定された HTTP コンテキストで使用できるかどうかを検証します。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="spContext">SharePointContext</param>
		/// <returns>使用できる場合は true を返します。</returns>
		protected override bool ValidateSharePointContext(HttpContextBase httpContext, SharePointContext spContext) {
			var spAcsContext = spContext as SharePointAcsContext;
			if (spAcsContext == null) {
				return false;
			}

			var spHostUrl = httpContext.Request.GetSPHostUrl();
			var contextToken = TokenHelper.GetContextTokenFromRequest(httpContext.Request);
			var spCacheKeyCookie = httpContext.Request.Cookies[SPCacheKeyKey];
			var spCacheKey = spCacheKeyCookie?.Value;

			return (
				spHostUrl == spAcsContext.SPHostUrl
				&& !string.IsNullOrEmpty(spAcsContext.CacheKey)
				&& spCacheKey == spAcsContext.CacheKey
				&& !string.IsNullOrEmpty(spAcsContext.ContextToken)
				&& (string.IsNullOrEmpty(contextToken)
					|| contextToken == spAcsContext.ContextToken)
			);
		}

		/// <summary>
		/// 指定された HTTP コンテキストに関連付けられた SharePointContext インスタンスを読み込みます。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <returns>SharePointContext インスタンスを返します。
		/// 見つからなければ <c>null</c> を返します。</returns>
		protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
			=> httpContext.Session[SPContextKey] as SharePointAcsContext;

		/// <summary>
		/// 指定された HTTP コンテキストに関連付けられている指定された SharePointContext インスタンスを保存します。
		/// <c>null</c>は、HTTP コンテキストに関連付けられた SharePointContext インスタンスを消去するために受け入れられます。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="spContext">保存する SharePointContext インスタンス</param>
		protected override void SaveSharePointContext(HttpContextBase httpContext, SharePointContext spContext) {
			var spAcsContext = spContext as SharePointAcsContext;

			if (spAcsContext != null) {
				HttpCookie spCacheKeyCookie = new HttpCookie(SPCacheKeyKey) {
					Value = spAcsContext.CacheKey,
					Secure = true,
					HttpOnly = true
				};

				httpContext.Response.AppendCookie(spCacheKeyCookie);
			}

			httpContext.Session[SPContextKey] = spAcsContext;
		}

		#endregion
	}
}