namespace EPaymentVoucher.Locators.MyTaskList
{
    public static class MTPaymentVoucherXPath
    {
        public const string TabMTPV = "/Tasks/MyTaskList#MTPV";
        public const string RequestNo = "//input[@id='RequestNo']";
        public const string RequestDateFrom = "//input[@id='RequestDate1']";
        public const string RequestDateTo = "//input[@id='RequestDate2']";
        public const string PaymentType = "//select[@id='PaymentType']";
        public const string PaymentTypeOption = "//select[@id='PaymentType']/option[@label='{x}']";
        public const string Purpose = "//select[@id='Purpose']";
        public const string PurposeOption = "//select[@id='Purpose']/option[@label='{x}']";
        public const string Payee = "//input[@id='Payee']";
        public const string Status = "//select[@id='Status']";
        public const string StatusOption = "//select[@id='Status']/option[@label='{x}']";
        public const string BtnSearch = "//div[@id='btnSearch']";
        public const string BtnClear = "//div[@id='btnClear']";
        public const string BtnSubmitAction = "//div[@id='btnDoAction']";
		public const string ListRowRequestNo = "//a[contains(@href,'ViewPVApproval')]";
		public const string ListRowAction = "//select[contains(@ng-options,'AvailableAction')]";
		public const string ListRowReturnTo = "//select[contains(@ng-model,'SelectedReturnTo')]";
		public const string ListRowNotes = "//input[contains(@ng-model,'ActionNotes')]";
		public const string Grid = "//div[@id='PVMyTaskListAngularAppDiv']//div[@role='presentation']";
        public const string RowStatus = "//div[contains(@ng-class, 'isSelected')]/div/div[7]/div";
	}
}
