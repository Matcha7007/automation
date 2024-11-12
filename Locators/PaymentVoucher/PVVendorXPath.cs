namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVVendorXPath
	{
		public const string BtnPayee = "//dirpopupwindow[contains(@url, 'Payee')]//a[@ng-click='open()']";
		public const string IsRecurringTrue = "//input[@name='recurring'][@ng-value='true']";
		public const string IsRecurringFalse = "//input[@name='recurring'][@ng-value='false']";
		public const string PO = "//select[@name='PO']";
		public const string GR = "//select[@name='GR']";
		public const string BtnAdd = "//div[@ng-click='getDataDetailByGR()']";
		public const string BtnDetail = "//dirpopupwindow[contains(@url, 'VendorPaymentVoucher')]//a[@ng-click='open()']";
		public const string ListRowDeleteDetail = "//tr[contains(@ng-repeat, 'DetailTransactions')]//a[@ng-click='DeleteItem($index)']";
		public const string ListRowEditDetail = "//tr[contains(@ng-repeat, 'DetailTransactions')]//dirpopupwindow[@initial-value='EditItem($index)']//a[@ng-click='open()']";
	}
}
