using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Models {
	public class AccessToken : Tuple<string, DateTime> {
		public AccessToken(string item1, DateTime item2) : base(item1, item2) {
		}

		/// <summary>
		/// Determines if the specified access token is valid.
		/// It considers an access token as not valid if it is null, or it has expired.
		/// </summary>
		/// <returns>True if the access token is valid.</returns>
		public bool IsValid
			=> !string.IsNullOrEmpty(this.Item1)
			&& this.Item2 > DateTime.UtcNow;
	}
}