namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb.Enums {
	/// <summary>
	/// リダイレクションの状態を表す列挙体です。
	/// </summary>
	public enum RedirectionStatus {
		/// <summary>
		/// OK
		/// </summary>
		Ok,

		/// <summary>
		/// リダイレクトする必要があります
		/// </summary>
		ShouldRedirect,

		/// <summary>
		/// リダイレクトできません
		/// </summary>
		CanNotRedirect
	}
}