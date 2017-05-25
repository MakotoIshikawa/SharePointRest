// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Enums;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Primitive;
using SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Token;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb {
	public partial class Default : Page {
		#region フィールド

		private SharePointContextToken _contextToken = null;

		/// <summary>
		/// Create a namespace manager for parsing the ATOM XML returned by the queries.
		/// </summary>
		private XmlNamespaceManager _xmlnspm = new XmlNamespaceManager(new NameTable());

		#endregion

		#region プロパティ

		/// <summary>
		/// アクセストークン
		/// </summary>
		protected string AccessToken { get; set; } = null;

		/// <summary>
		/// SharePoint Uri
		/// </summary>
		protected Uri SharepointUrl { get; set; } = null;

		#endregion

		#region イベントハンドラ

		protected void AddList_Click(object sender, EventArgs e) {
			var commandAccessToken = ((Button)sender).CommandArgument;
			if (this.AddListNameBox.Text != "") {
				this.AddList(commandAccessToken, AddListNameBox.Text);
			} else {
				this.AddListNameBox.Text = "Enter a list title";
			}
		}

		protected void RefreshList_Click(object sender, EventArgs e) {
			var commandAccessToken = (sender as Button)?.CommandArgument;
			this.RetrieveLists(commandAccessToken);
		}

		protected void RetrieveListButton_Click(object sender, EventArgs e) {
			var commandAccessToken = ((Button)sender).CommandArgument;
			var listId = new Guid();
			if (Guid.TryParse(RetrieveListNameBox.Text, out listId)) {
				this.RetrieveListItems(commandAccessToken, listId);
			} else {
				this.RetrieveListNameBox.Text = "Enter a List GUID";
			}
		}

		protected void AddItemButton_Click(object sender, EventArgs e) {
			var commandAccessToken = ((Button)sender).CommandArgument;
			var listId = new Guid(RetrieveListNameBox.Text);
			if (this.AddListItemBox.Text != "") {
				this.AddListItem(commandAccessToken, listId, AddListItemBox.Text);
			} else {
				this.AddListItemBox.Text = "Enter an item title";
			}
		}

		protected void DeleteListButton_Click(object sender, EventArgs e) {
			var commandAccessToken = ((Button)sender).CommandArgument;
			var listId = new Guid(RetrieveListNameBox.Text);
			this.DeleteList(commandAccessToken, listId);
		}

		protected void ChangeListTitleButton_Click(object sender, EventArgs e) {
			var commandAccessToken = ((Button)sender).CommandArgument;
			var listId = new Guid(RetrieveListNameBox.Text);
			if (!string.IsNullOrEmpty(ChangeListTitleBox.Text)) {
				this.ChangeListTitle(commandAccessToken, listId, ChangeListTitleBox.Text);
			} else {
				this.ChangeListTitleBox.Text = "Enter a new list title";
			}
		}

		#endregion

		#region メソッド

		/// <summary>
		/// ページ初期化時に呼ばれます。
		/// </summary>
		/// <param name="sender">送信元</param>
		/// <param name="e">イベントデータ</param>
		protected void Page_PreInit(object sender, EventArgs e) {
			Uri redirectUrl;
			var st = SharePointContextProvider.CheckRedirectionStatus(this.Context, out redirectUrl);
			switch (st) {
			case RedirectionStatus.ShouldRedirect:
				Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
				break;
			case RedirectionStatus.CanNotRedirect:
				Response.Write("An error occurred while processing your request.");
				Response.End();
				break;
			}
		}

		/// <summary>
		/// Page_load メソッドは、コンテキストトークンとアクセストークンを取得します。
		/// アクセストークンは、すべてのデータ取得メソッドで使用されます。
		/// </summary>
		/// <param name="sender">送信元</param>
		/// <param name="e">イベントデータ</param>
		protected void Page_Load(object sender, EventArgs e) {
			string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

			if (contextTokenString != null) {
				_contextToken =
					TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);

				this.SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
				this.AccessToken = TokenHelper.GetAccessToken(_contextToken, SharepointUrl.Authority).AccessToken;

				// In a production add-in, you should cache the access token somewhere, such as in a database
				// or ASP.NET Session Cache. (Do not put it in a cookie.) Your code should also check to see 
				// if it is expired before using it (and use the refresh token to get a new one when needed). 
				// For more information, see the MSDN topic at https://msdn.microsoft.com/library/office/dn762763.aspx
				// For simplicity, this sample does not follow these practices. 
				this.AddListButton.CommandArgument = AccessToken;
				this.RefreshListButton.CommandArgument = AccessToken;
				this.RetrieveListButton.CommandArgument = AccessToken;
				this.AddItemButton.CommandArgument = AccessToken;
				this.DeleteListButton.CommandArgument = AccessToken;
				this.ChangeListTitleButton.CommandArgument = AccessToken;
				this.RetrieveLists(AccessToken);

			} else if (!IsPostBack) {
				this.Response.Write("Could not find a context token.");
			}
		}

		private void RetrieveLists(string accessToken) {
			if (IsPostBack) {
				SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
			}

			AddItemButton.Visible = false;
			AddListItemBox.Visible = false;
			DeleteListButton.Visible = false;
			ChangeListTitleButton.Visible = false;
			ChangeListTitleBox.Visible = false;
			RetrieveListNameBox.Enabled = true;
			ListTable.Rows[0].Cells[1].Text = "List ID";

			//Add needed namespaces to the namespace manager.
			_xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
			_xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
			_xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");

			//Execute a REST request for all of the site's lists.
			HttpWebRequest listRequest =
				(HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists");
			listRequest.Method = "GET";
			listRequest.Accept = "application/atom+xml";
			listRequest.ContentType = "application/atom+xml;type=entry";
			listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
			HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
			StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
			var listXml = new XmlDocument();
			listXml.LoadXml(listReader.ReadToEnd());

			var titleList = listXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", _xmlnspm);
			var idList = listXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Id", _xmlnspm);

			int listCounter = 0;
			foreach (XmlNode title in titleList) {
				TableRow tableRow = new TableRow();
				LiteralControl idClick = new LiteralControl();
				//Use Javascript to populate the RetrieveListNameBox control with the list id.
				string clickScript = "<a onclick=\"document.getElementById(\'RetrieveListNameBox\').value = '" + idList[listCounter].InnerXml + "';\" href=\"#\">" + idList[listCounter].InnerXml + "</a>";
				idClick.Text = clickScript;
				TableCell tableCell1 = new TableCell();
				tableCell1.Controls.Add(new LiteralControl(title.InnerXml));
				TableCell tableCell2 = new TableCell();
				tableCell2.Controls.Add(idClick);
				tableRow.Cells.Add(tableCell1);
				tableRow.Cells.Add(tableCell2);
				ListTable.Rows.Add(tableRow);
				listCounter++;
			}
		}

		private void RetrieveListItems(string accessToken, Guid listId) {
			if (IsPostBack) {
				SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
			}

			//Adjust the visibility of controls on the page in light of the list-specific context.
			AddItemButton.Visible = true;
			AddListItemBox.Visible = true;
			DeleteListButton.Visible = true;
			ChangeListTitleButton.Visible = true;
			ChangeListTitleBox.Visible = true;
			RetrieveListNameBox.Enabled = false;
			ListTable.Rows[0].Cells[1].Text = "List Items";

			//Add needed namespaces to the namespace manager.
			_xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
			_xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
			_xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");

			//Execute a REST request to get the list name.
			HttpWebRequest listRequest =
				(HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'" + listId + "')");
			listRequest.Method = "GET";
			listRequest.Accept = "application/atom+xml";
			listRequest.ContentType = "application/atom+xml;type=entry";
			listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
			HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
			StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
			var listXml = new XmlDocument();
			listXml.LoadXml(listReader.ReadToEnd());

			var listNameNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:Title", _xmlnspm);
			string listName = listNameNode.InnerXml;

			//Execute a REST request to get all of the list's items.
			HttpWebRequest itemRequest =
				(HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'" + listId + "')/Items");
			itemRequest.Method = "GET";
			itemRequest.Accept = "application/atom+xml";
			itemRequest.ContentType = "application/atom+xml;type=entry";
			itemRequest.Headers.Add("Authorization", "Bearer " + accessToken);
			HttpWebResponse itemResponse = (HttpWebResponse)itemRequest.GetResponse();
			StreamReader itemReader = new StreamReader(itemResponse.GetResponseStream());
			var itemXml = new XmlDocument();
			itemXml.LoadXml(itemReader.ReadToEnd());

			var itemList = itemXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", _xmlnspm);

			TableRow tableRow = new TableRow();
			TableCell tableCell1 = new TableCell();
			tableCell1.Controls.Add(new LiteralControl(listName));
			TableCell tableCell2 = new TableCell();

			foreach (XmlNode itemTitle in itemList) {
				tableCell2.Text += itemTitle.InnerXml + "<br>";
			}

			tableRow.Cells.Add(tableCell1);
			tableRow.Cells.Add(tableCell2);
			ListTable.Rows.Add(tableRow);
		}

		private void AddList(string accessToken, string newListName) {
			if (this.IsPostBack) {
				this.SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
			}

			try {
				//Add pertinent namespace to the namespace manager.
				this._xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");

				var formDigest = string.Empty;
				{
					//Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
					//Web require the form digest in the request header.
					var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/contextinfo");
					request.Method = "POST";
					request.ContentType = "text/xml;charset=utf-8";
					request.ContentLength = 0;
					request.Headers.Add("Authorization", "Bearer " + accessToken);

					var response = (HttpWebResponse)request.GetResponse();

					using (var stream = response.GetResponseStream())
					using (var reader = new StreamReader(stream, Encoding.UTF8)) {
						var xmlDoc = new XmlDocument();
						xmlDoc.LoadXml(reader.ReadToEnd());

						var node = xmlDoc.SelectSingleNode("//d:FormDigestValue", _xmlnspm);
						formDigest = node.InnerXml;
					}
				}
				{
					//Execute a REST request to add a list that has the user-supplied name.
					//The body of the REST request is ASCII encoded and inserted into the request stream.
					var listPostBody = $"{{'__metadata':{{'type':'SP.List'}}, 'Title':'{newListName}', 'BaseTemplate': 100}}";
					var listPostData = Encoding.ASCII.GetBytes(listPostBody);

					var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/lists");
					request.Method = "POST";
					request.ContentLength = listPostBody.Length;
					request.ContentType = "application/json;odata=verbose";
					request.Accept = "application/json;odata=verbose";
					request.Headers.Add("Authorization", $"Bearer {accessToken}");
					request.Headers.Add("X-RequestDigest", formDigest);

					using (var stream = request.GetRequestStream()) {
						stream.Write(listPostData, 0, listPostData.Length);
						stream.Close();
					}

					var response = (HttpWebResponse)request.GetResponse();
				}

				this.RetrieveLists(accessToken);
			} catch (Exception e) {
				this.AddListNameBox.Text = e.Message;
			}
		}

		private void AddListItem(string accessToken, Guid listId, string newItemName) {
			if (this.IsPostBack) {
				this.SharepointUrl = new Uri(this.Request.QueryString["SPHostUrl"]);
			}

			try {
				//Add pertinent namespaces to the namespace manager.
				this._xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
				this._xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
				this._xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");


				//Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
				//Web require the form digest in the request header.
				var contextinfoRequest = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/contextinfo");
				contextinfoRequest.Method = "POST";
				contextinfoRequest.ContentType = "text/xml;charset=utf-8";
				contextinfoRequest.ContentLength = 0;
				contextinfoRequest.Headers.Add("Authorization", "Bearer " + accessToken);

				var contextinfoResponse = (HttpWebResponse)contextinfoRequest.GetResponse();
				var contextinfoReader = new StreamReader(contextinfoResponse.GetResponseStream(), Encoding.UTF8);
				var formDigestXML = new XmlDocument();
				formDigestXML.LoadXml(contextinfoReader.ReadToEnd());
				var formDigestNode = formDigestXML.SelectSingleNode("//d:FormDigestValue", _xmlnspm);
				var formDigest = formDigestNode.InnerXml;

				//Execute a REST request to get the list name and the entity type name for the list.
				var listRequest = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'" + listId + "')");
				listRequest.Method = "GET";
				listRequest.Accept = "application/atom+xml";
				listRequest.ContentType = "application/atom+xml;type=entry";
				listRequest.Headers.Add("Authorization", "Bearer " + accessToken);

				var listResponse = (HttpWebResponse)listRequest.GetResponse();
				var listReader = new StreamReader(listResponse.GetResponseStream());
				var listXml = new XmlDocument();
				listXml.LoadXml(listReader.ReadToEnd());

				//The entity type name is the required type when you construct a request to add a list item.
				var entityTypeNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:ListItemEntityTypeFullName", _xmlnspm);
				var listNameNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:Title", _xmlnspm);
				var entityTypeName = entityTypeNode.InnerXml;
				var listName = listNameNode.InnerXml;

				//Execute a REST request to add an item to the list.
				var itemPostBody = "{'__metadata':{'type':'" + entityTypeName + "'}, 'Title':'" + newItemName + "'}}";
				var itemPostData = Encoding.ASCII.GetBytes(itemPostBody);

				var itemRequest = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'" + listId + "')/Items");
				itemRequest.Method = "POST";
				itemRequest.ContentLength = itemPostBody.Length;
				itemRequest.ContentType = "application/json;odata=verbose";
				itemRequest.Accept = "application/json;odata=verbose";
				itemRequest.Headers.Add("Authorization", "Bearer " + accessToken);
				itemRequest.Headers.Add("X-RequestDigest", formDigest);

				var itemRequestStream = itemRequest.GetRequestStream();

				itemRequestStream.Write(itemPostData, 0, itemPostData.Length);
				itemRequestStream.Close();

				var itemResponse = (HttpWebResponse)itemRequest.GetResponse();

				this.RetrieveListItems(accessToken, listId);
			} catch (Exception e) {
				this.AddListItemBox.Text = e.Message;
			}
		}

		private void ChangeListTitle(string accessToken, Guid listId, string newListTitle) {
			if (this.IsPostBack) {
				this.SharepointUrl = new Uri(this.Request.QueryString["SPHostUrl"]);
			}

			//Add pertinent namespace to the namespace manager.
			this._xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");

			var formDigest = string.Empty;
			{
				//Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
				//Web require the form digest in the request header.
				var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/contextinfo");
				request.Method = "POST";
				request.ContentType = "text/xml;charset=utf-8";
				request.ContentLength = 0;
				request.Headers.Add("Authorization", "Bearer " + accessToken);

				var response = (HttpWebResponse)request.GetResponse();

				using (var stream = response.GetResponseStream())
				using (var reader = new StreamReader(stream, Encoding.UTF8)) {
					var xmlDoc = new XmlDocument();
					xmlDoc.LoadXml(reader.ReadToEnd());
					var node = xmlDoc.SelectSingleNode("//d:FormDigestValue", _xmlnspm);
					formDigest = node.InnerXml;
				}
			}

			var eTag = string.Empty;
			{
				//Execute a REST request to get the ETag value, which needs to be sent with the delete request.
				var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'" + listId + "')");
				request.Method = "GET";
				request.Accept = "application/atom+xml";
				request.ContentType = "application/atom+xml;type=entry";
				request.Headers.Add("Authorization", "Bearer " + accessToken);

				var response = (HttpWebResponse)request.GetResponse();
				eTag = response.Headers["ETag"];
			}
			{
				//Execute a REST request to change the list title
				var body = "{'__metadata':{'type':'SP.List'}, 'Title':'" + newListTitle + "'}";
				var data = Encoding.ASCII.GetBytes(body);

				var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/lists(guid'" + listId + "')");
				request.Method = "POST";
				request.ContentLength = body.Length;
				request.ContentType = "application/json;odata=verbose";
				request.Accept = "application/json;odata=verbose";
				request.Headers.Add("Authorization", "Bearer " + accessToken);
				request.Headers.Add("X-RequestDigest", formDigest);
				request.Headers.Add("If-Match", eTag);
				request.Headers.Add("X-Http-Method", "MERGE");

				using (var stream = request.GetRequestStream()) {
					stream.Write(data, 0, data.Length);
					stream.Close();
				}

				var response = (HttpWebResponse)request.GetResponse();
			}

			RetrieveListItems(accessToken, listId);
		}

		private void DeleteList(string accessToken, Guid listId) {
			if (this.IsPostBack) {
				this.SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
			}

			//Add pertinent namespace to the namespace manager.
			this._xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");

			var formDigest = string.Empty;
			{
				//Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
				//Web require the form digest in the request header.
				var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/contextinfo");
				request.Method = "POST";
				request.ContentType = "text/xml;charset=utf-8";
				request.ContentLength = 0;
				request.Headers.Add("Authorization", "Bearer " + accessToken);

				var response = (HttpWebResponse)request.GetResponse();

				using (var stream = response.GetResponseStream())
				using (var reader = new StreamReader(stream, Encoding.UTF8)) {
					var xmlDoc = new XmlDocument();
					xmlDoc.LoadXml(reader.ReadToEnd());

					var node = xmlDoc.SelectSingleNode("//d:FormDigestValue", _xmlnspm);
					formDigest = node.InnerXml;
				}
			}

			var eTag = string.Empty;
			{
				//Execute a REST request to get the ETag value, which needs to be sent with the delete request.
				var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'{listId}')");
				request.Method = "GET";
				request.Accept = "application/atom+xml";
				request.ContentType = "application/atom+xml;type=entry";
				request.Headers.Add("Authorization", "Bearer " + accessToken);

				var response = (HttpWebResponse)request.GetResponse();
				eTag = response.Headers["ETag"];
			}
			{
				//Execute a REST request to delete the list.
				var request = (HttpWebRequest)WebRequest.Create($"{this.SharepointUrl}/_api/Web/lists(guid'{listId}')");
				request.Method = "POST";
				request.ContentLength = 0;
				request.ContentType = "text/xml;charset=utf-8";
				request.Headers.Add("X-RequestDigest", formDigest);
				request.Headers.Add("If-Match", eTag);
				request.Headers.Add("Authorization", "Bearer " + accessToken);
				request.Headers.Add("X-Http-Method", "DELETE");

				var response = (HttpWebResponse)request.GetResponse();
			}

			this.RetrieveListNameBox.Text = string.Empty;
			this.RetrieveLists(accessToken);
		}

		#endregion
	}
}

/*
SharePoint Add-in REST/OData Basic Data Operations, https://github.com/OfficeDev/SharePoint-Add-in-REST-OData-BasicDataOperations
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/
