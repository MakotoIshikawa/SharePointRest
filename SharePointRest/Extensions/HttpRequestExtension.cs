using System;
using System.Web;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;
using static SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Const.ConstString;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Extensions {
	/// <summary>
	/// HttpRequest を拡張するメソッドを提供します。
	/// </summary>
	public static partial class HttpRequestExtension {
		#region メソッド

		/// <summary>
		/// 指定された HTTP 要求の QueryString から SharePoint ホストURLを取得します。
		/// </summary>
		/// <param name="httpRequest">HTTP 要求</param>
		/// <returns>SharePoint ホストの URL。
		/// HTTP 要求に SharePoint ホストの URL が含まれていない場合は、
		/// <c>null</ c> を返します。</returns>
		public static Uri GetSPHostUrl(this HttpRequestBase httpRequest) {
			if (httpRequest == null) {
				throw new ArgumentNullException(nameof(httpRequest));
			}

			var spHostUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SPHostUrlKey]);
			Uri spHostUrl;
			if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl)
			&& (spHostUrl.Scheme == Uri.UriSchemeHttp
				|| spHostUrl.Scheme == Uri.UriSchemeHttps
			)) {
				return spHostUrl;
			}

			return null;
		}

		/// <summary>
		/// 指定された HTTP 要求の QueryString から SharePoint ホストURLを取得します。
		/// </summary>
		/// <param name="httpRequest">HTTP 要求</param>
		/// <returns>SharePoint ホストの URL。
		/// HTTP 要求に SharePoint ホストの URL が含まれていない場合は、
		/// <c>null</ c> を返します。</returns>
		public static Uri GetSPHostUrl(this HttpRequest httpRequest)
			=> GetSPHostUrl(new HttpRequestWrapper(httpRequest));

		#endregion
	}
}