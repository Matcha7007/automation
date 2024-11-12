namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVInternalBankXPath
	{
		public const string BtnAddDetail = "//dirpopupwindow[contains(@url, 'InternalBank')]//a[@ng-click='open()']";
		public const string ListBtnDelete = "//a[contains(@ng-click,'DeleteItem')]";
		public const string ListRowEditDetail = "//dirpopupwindow[@initial-value='EditItem($index)']//a[@ng-click='open()']";
		public const string FromBankAccountNo = "//select[@id='FromBankAccountNo']";
		public const string ToBankAccountNo = "//select[@id='ToBankAccountNo']";
		public const string CostCenter = "//select[@id='CostCenter']";
		public const string ShortDescription = "//input[@id='ShortDescription']";
		public const string LongDescription = "//textarea[@id='LongDescription']";
		public const string Currency = "//select[@id='Currency']";
		public const string Amount = "//input[@id='Amount']";
		public const string BtnAttachment = $"//span[contains(@class, 'custom-upload')]";
		public const string Remarks = "//textarea[@id='Remarks']";
		public const string BtnSavePopUp = $"//div[@ng-click='OK()']";
	}
}
