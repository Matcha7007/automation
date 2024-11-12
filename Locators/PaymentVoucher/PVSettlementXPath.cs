namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVSettlementXPath
	{
		public const string BtnRequestorName = "//dirpopupwindow[contains(@url, 'RequestorInformation')]//a[@ng-click='open()']";
		public const string PaymentVoucherNo = "//select[@ng-model='selectedPaymentVoucher']";
		public const string BtnAddDetail = "//dirpopupwindow[contains(@ng-show, 'Purpose==')]//a[@ng-click='open()']";
		public const string BtnAddDetail2 = "//dirpopupwindow[contains(@ng-show, 'Purpose!=')]//a[@ng-click='open()']";
		public const string BtnBusinessName = "//div[@id=' vendor-business-name']";
		public const string CostCenterName = "//select[@ng-model='SelectedCostCenter']";
		public const string ShortDescription = "//input[@id='ShortDescription']";
		public const string LongDescription = "//textarea[@id='LongDescription']";
		public const string Date = "//input[@id='Date']";
		public const string Amount = "//input[@id='Amount']";
		public const string BtnFile = $"//div[@class='row'][not(@ng-show)]{PVHeaderXPath.BaseBtnFile}";
		public const string BtnSavePopUp = $"//div[@ng-click='OK(popUpform)']";
	}
}
