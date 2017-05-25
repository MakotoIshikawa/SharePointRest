using System;
using System.Security.Principal;
using System.Web;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Extensions;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive;
using static SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Const.ConstString;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.HighTrust {
	/// <summary>
	/// Default provider for SharePointHighTrustContext.
	/// </summary>
	public class SharePointHighTrustContextProvider : SharePointContextProvider {
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
			var logonUserIdentity = httpRequest.LogonUserIdentity;
			if (logonUserIdentity == null || !logonUserIdentity.IsAuthenticated || logonUserIdentity.IsGuest || logonUserIdentity.User == null) {
				return null;
			}

			return new SharePointHighTrustContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, logonUserIdentity);
		}

		/// <summary>
		/// 指定された SharePoint コンテキストを指定された HTTP コンテキストで使用できるかどうかを検証します。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="spContext">SharePointContext</param>
		/// <returns>使用できる場合は true を返します。</returns>
		protected override bool ValidateSharePointContext(HttpContextBase httpContext, SharePointContext spContext) {
			var spHighTrustContext = spContext as SharePointHighTrustContext;

			if (spHighTrustContext != null) {
				Uri spHostUrl = httpContext.Request.GetSPHostUrl();
				WindowsIdentity logonUserIdentity = httpContext.Request.LogonUserIdentity;

				return spHostUrl == spHighTrustContext.SPHostUrl &&
					   logonUserIdentity != null &&
					   logonUserIdentity.IsAuthenticated &&
					   !logonUserIdentity.IsGuest &&
					   logonUserIdentity.User == spHighTrustContext.LogonUserIdentity.User;
			}

			return false;
		}

		/// <summary>
		/// 指定された HTTP コンテキストに関連付けられた SharePointContext インスタンスを読み込みます。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <returns>SharePointContext インスタンスを返します。
		/// 見つからなければ <c>null</c> を返します。</returns>
		protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
			=> httpContext.Session[SPContextKey] as SharePointHighTrustContext;

		/// <summary>
		/// 指定された HTTP コンテキストに関連付けられている指定された SharePointContext インスタンスを保存します。
		/// <c>null</c>は、HTTP コンテキストに関連付けられた SharePointContext インスタンスを消去するために受け入れられます。
		/// </summary>
		/// <param name="httpContext">HTTP コンテキスト</param>
		/// <param name="spContext">保存する SharePointContext インスタンス</param>
		protected override void SaveSharePointContext(HttpContextBase httpContext, SharePointContext spContext)
			=> httpContext.Session[SPContextKey] = spContext as SharePointHighTrustContext;

		#endregion
	}
}