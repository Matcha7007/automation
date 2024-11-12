namespace EPaymentVoucher.Locators.MyTaskList
{
    public static class MTVendorXPath
    {
        public const string TabMTVD = "//a[@href='#MTVD']";
        public const string MTVDVendorCode = "//input[@id='VendorCode']";
        public const string MTVDVendorName = "//input[@id='VendorName']";
        public const string MTVDBusinessName = "//input[@id='BusinessName']";
        public const string MTVDVBtnSearch = "//a[@id='VendorSearchBtn']";
        public const string MTVDVBtnClear = "//div[@id='VendorClearBtn']";
		public const string RowsRequestNo = "//a[contains(@href, 'Vendor/DoAction')]";
		public const string RowsRequestNoResubmit = "//a[contains(@href, 'Vendor/Edit')]";
		public const string RowVendorCode = "//tr[1]//a[contains(@href, 'Vendor/Display')]";
		public const string RowVendorName = "//table[@id='vendortable']//tr[1]/td[3]";
		public const string Action = "//select[@id='ApprovalAction']";
		public const string ReturnTo = "//select[@id='ReturnTo']";
		public const string Notes = "//textarea[@id='ApprovalAction_AppAct_Notes']";
		public const string BtnSubmitAction = "//div[@id='submit-btn']";
		public const string BtnSubmitActionResubmit = "//div[@id='submitting']";
		public const string RowRole = "//tbody[@id='TableActionHistory']/tr[1]/td[2]";
		public const string RowAction = "//tbody[@id='TableActionHistory']/tr[1]/td[4]";
		public const string TabProfile = "//a[@href='#vp']";
		public const string TabApproval = "//a[@href='#aa']";
		public const string FieldVendorCode = "//input[@id='VendorCode']";
		public const string AlertError = "//div[contains(@class, 'alert-danger')]";
	}
}
