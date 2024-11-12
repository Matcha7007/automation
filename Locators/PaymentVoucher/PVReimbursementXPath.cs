namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVReimbursementXPath
	{
		public const string BtnAddListRequestorName = "//dirpopupwindow[@ng-model='selectedRequestor']//a[@ng-click='open()']";
		public const string BankName = "//select[@id='BankName']";
		public const string RequestorBankAccountName = "//input[@name='RequestorBankAccountName']";
		public const string RequestorBankAccountNo = "//input[@name='RequestorBankAccoutnNo']";
		public const string BtnAddDetail = "//dirpopupwindow[contains(@url, 'ReimbursementMainPopUp')][@initial-value='NewItem()']//a[@ng-click='open()']";
		public const string BtnAddDetail2 = "//dirpopupwindow[contains(@url, 'OtherPopUp')][@initial-value='NewItem()']//a[@ng-click='open()']";
		public const string ListRowDeleteDetail = "//tr[contains(@ng-repeat, 'ListActivityFormData')]//a[@ng-click='DeleteItem($index)']";
	}
}
