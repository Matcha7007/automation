namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVAdvanceXPath
	{
		public const string BtnRequestorName = "//dirpopupwindow[contains(@url, 'RequestorInformation')]//a[@ng-click='open()']";
		public const string BankName = "//select[@id='BankName']";
		public const string RequestorBankAccountName = "//input[@name='RequestorBankAccountName']";
		public const string RequestorBankAccountNo = "//input[@name='RequestorBankAccoutnNo']";
		public const string StartDate = "//input[@id='StartDate']";
		public const string EndDate = "//input[@id='EndDate']";
		public const string Amount = "//input[@name='Amount']";
		public const string Reason = "//input[@name='Reason']";
	}
}
