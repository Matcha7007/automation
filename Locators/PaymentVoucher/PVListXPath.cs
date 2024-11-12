namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVListXPath
	{
		public const string UrlPaymentVoucher = "/PaymentVoucher";
		public const string UrlCreatePaymentList = $"{UrlPaymentVoucher}/Create";

		#region List Page
		public const string RequestNo = "//input[@id='RequestNo']";
		public const string PVNo = "//input[@id='PVNo']";
		public const string RequestDateFrom = "//input[@id='RequestDate1']";
		public const string RequestDateTo = "//input[@id='RequestDate2']";
		public const string PaymentType = "//select[@id='PaymentType']";
		public const string Purpose = "//select[@id='Purpose']";
		public const string Payee = "//input[@id='Payee']";
		public const string Title = "//input[@id='Title']";
		public const string Status = "//select[@id='Status']";
		public const string BtnSearch = "//div[@id='btnSearch']";
		public const string BtnClear = "//div[@id='btnClear']";
		public const string BtnCreate = $"//a[@href='{UrlCreatePaymentList}']";
		public const string RowRequestNo = "//a[contains(@href, 'PaymentVoucher/edit')]/span";
		#endregion

		#region Header Create
		public const string HeaderPaymentType = "//select[@name='PaymentType']";
		public const string HeaderPaymentPurpose = "//select[@ng-model='$parent.Purposes']";
		public const string HeaderBtnNext = "//div[contains(@class, 'btn-next')]";
		public const string HeaderBtnCancel = "//div[contains(@class, 'btn-cancel')]";
		#endregion
	}
}
