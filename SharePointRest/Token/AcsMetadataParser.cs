using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web.Script.Serialization;
using static SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token.TokenHelper;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token {
	/// <summary>
	/// This class is used to get MetaData document from the global STS endpoint. It contains
	/// methods to parse the MetaData document and get endpoints and STS certificate.
	/// </summary>
	public static class AcsMetadataParser {
		#region メソッド

		public static X509Certificate2 GetAcsSigningCert(string realm) {
			var document = GetMetadataDocument(realm);

			if (null != document.keys && document.keys.Count > 0) {
				var signingKey = document.keys[0];

				if (null != signingKey && null != signingKey.keyValue) {
					return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
				}
			}

			throw new Exception("Metadata document does not contain ACS signing certificate.");
		}

		public static string GetDelegationServiceUrl(string realm) {
			var document = GetMetadataDocument(realm);

			var delegationEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == DelegationIssuance);
			if (null != delegationEndpoint) {
				return delegationEndpoint.location;
			}

			throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
		}

		public static string GetStsUrl(string realm) {
			var document = GetMetadataDocument(realm);

			var s2sEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

			if (null != s2sEndpoint) {
				return s2sEndpoint.location;
			}

			throw new Exception("Metadata document does not contain STS endpoint url");
		}

		private static JsonMetadataDocument GetMetadataDocument(string realm) {
			string acsMetadataEndpointUrlWithRealm = string.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
																  TokenHelper.GetAcsMetadataEndpointUrl(),
																   realm);
			byte[] acsMetadata;
			using (WebClient webClient = new WebClient()) {

				acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
			}
			string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

			JavaScriptSerializer serializer = new JavaScriptSerializer();
			JsonMetadataDocument document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

			if (null == document) {
				throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
			}

			return document;
		}

		#endregion

		private class JsonMetadataDocument {
			public string serviceName { get; set; }
			public List<JsonEndpoint> endpoints { get; set; }
			public List<JsonKey> keys { get; set; }
		}

		private class JsonEndpoint {
			public string location { get; set; }
			public string protocol { get; set; }
			public string usage { get; set; }
		}

		private class JsonKeyValue {
			public string type { get; set; }
			public string value { get; set; }
		}

		private class JsonKey {
			public string usage { get; set; }
			public JsonKeyValue keyValue { get; set; }
		}
	}
}