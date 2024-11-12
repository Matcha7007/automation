namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVPolicyHolderXPath
	{
		public const string BtnAddListRequestorName = "//dirpopupwindow[@ng-model='selectedRequestor']//a[@ng-click='open()']";
		public const string ListBtnDelete = "//a[contains(@ng-click,'DeleteItem')]";
		public const string ListRowEditDetail = "//dirpopupwindow[@initial-value='EditItem($index)']//a[@ng-click='open()']";
		public const string BtnAddDetail = "//dirpopupwindow[contains(@url, 'PolicyHolder')]//a[@ng-click='open()']";
		public const string Payee  = "//input[@id='Payee']";
		public const string PolicyNo  = "//input[@id='PolicyNo']";
		public const string SPAJNo  = "//input[@id='SpajNo']";
		public const string BussinessName  = "//dirpopupwindow[contains(@url, 'PVBusinessName')]//a[@ng-click='open()']";
		public const string LongDesc  = "//textarea[@id='LongDescription']";
		public const string ShortDesc  = "//input[@id='ShortDescription']";
		public const string Product  = "//select[@ng-model='selectedProduct']";
		public const string Distribution  = "//select[@ng-model='selectedDistribution']";
		public const string Country  = "//select[@ng-model='selectedCountry']";
		public const string BankName  = "//select[@id='BankName']";
		public const string BankAccountNo  = "//input[@id='BankAccountNo']";
		public const string BankAccountName  = "//input[@name='BankAccountName']";
		public const string Branch  = "//input[@id='Branch']";
		public const string Currency  = "//select[@id='Currency']";
		public const string GrossAmount  = "//input[@id='GrossAmount']";
		public const string TaxAmount  = "//input[@id='TaxAmount']";
		public const string ChargesAdmin  = "//input[@id='ChargesAdmin']";
		public const string ChargesMedical = "//input[@id='ChargesMedical']";
		public const string ForPremiumPayment  = "//input[@id='ForPremiumPayment']";
		public const string BtnAttachment  = "//div[contains(@ng-class, '.Attachment.')]//span[contains(@class, 'custom-upload')]";
		public const string TaxIdNo  = "//input[contains(@ng-model, 'arrayOftaxID')]";
		public const string TaxIdName  = "//input[@id='TaxIdName']";
		public const string TaxIdAddress  = "//textarea[@id='TaxIdAddress']";
		public const string BtnTaxIdAtaachmentPath  = "//div[contains(@ng-class, 'TaxIdAttachment')]//span[contains(@class, 'custom-upload')]";
		public const string Remarks  = "//textarea[@id='Remarks']";
		public const string PopupSave = "//div[@ng-click='OK()']";
	}
}
