namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVSpecialCustomerXPath
	{
		public const string BtnAddListRequestorName = "//dirpopupwindow[@ng-model='selectedRequestor']//a[@ng-click='open()']";
		public const string ListBtnDelete = "//a[contains(@ng-click,'DeleteItem')]";
		public const string ListRowEditDetail = "//dirpopupwindow[@initial-value='EditItem($index)']//a[@ng-click='open()']";
		public const string BtnAddDetail = "//dirpopupwindow[contains(@url, 'SpecialCustomer')]//a[@ng-click='open()']";
	}
	public class PVSpecialCustomerDetailXPath
	{
		public const string Payee = "//input[@id='Payee']";
		public const string PayeeType = "//select[@id='PayeeType']";
		public const string Country = "//select[@id='Country']";
		public const string CoDTrue = $"//input[@id='CoDTrue']";
		public const string CoDFalse = $"//input[@id='CoDFalse']";
		public const string BtnCoDAttachment = $"//div[contains(@ng-class, 'CodAttachment')]//span[contains(@ng-hide, 'Kas Negara')]";
		public const string CoDExpDate = $"//input[@id='CodExpiryDate']";
		public const string BtnSPPKPAttachment = $"//div[contains(@ng-class, 'SPPKP')]//span[contains(@ng-hide, 'Kas Negara')]";
		public const string TaxableTrue = $"//input[contains(@ng-model, 'TaxableEntrepreneur')][@id='TaxTrue']";
		public const string TaxableFalse = $"//input[contains(@ng-model, 'TaxableEntrepreneur')][@id='TaxFalse']";
		public const string UWTaxTrue = $"//input[contains(@ng-model, 'UnderWithholdingTax')][@id='TaxTrue']";
		public const string UWTaxFalse = $"//input[contains(@ng-model, 'UnderWithholdingTax')][@id='TaxFalse']";
		public const string TaxID = $"//input[contains(@ng-model,'arrayOftax')]";
		public const string TaxIDName = $"//input[@id='TaxIdName']";
		public const string TaxIDAddress = $"//textarea[@id='TaxIdAddress']";
		public const string BtnTaxIDAttachment = $"//div[contains(@ng-class, 'TaxIdAttachment')]//span[contains(@ng-hide, 'Kas Negara')]";
		public const string BankName = "//input[@id='TxtBankName']";
		public const string BankNameSelect = "//select[@id='BankName']";
		public const string BankAccountNo = $"//input[@id='BankAccountNo']";
		public const string BankAccountName = $"//input[@id='BankAccountName']";
		public const string Branch = $"//input[@id='Branch']";
		public const string BtnAddDetail = $"//dirpopupwindow[@initial-value='NewItemDetailList()']/a[@ng-click='open()']";
		public const string BtnAddDetail2 = $"//dirpopupwindow[@initial-value='NewItem()']/a[@ng-click='open()']";
		public const string BtnSavePopUp = $"//div[@ng-click='OK()']";
	}
	public class PVSpecialCustomerDetailDetailXPath
	{
		public const string BtnBusinessName = "//dirpopupwindow[contains(@url, 'PVBusinessName')]//a[@ng-click='open()']";
		public const string CostCenterName = "//select[@ng-model='selectedCostCenter']";
		public const string ShortDescription = "//input[@id='ShortDescription']";
		public const string LongDescription = "//textarea[@id='LongDescription']";
		public const string GrossAmount = "//input[@id='GrossAmount']";
		public const string Currency = "//select[@id='Currency']";
		public const string TaxExTrue = $"//input[@id='CoDTrue'][contains(@ng-model, 'TaxExemption')]";
		public const string TaxExFalse = $"//input[@id='CoDFalse'][contains(@ng-model, 'TaxExemption')]";
		public const string BtnSKBAttachment = $"//div[contains(@ng-class, 'SKBAttachment')]//span[contains(@class, 'custom-upload')]";
		public const string BtnAttachment = $"//div[@class='row'][9]//span[contains(@class, 'custom-upload')]";
		public const string Remarks = "//textarea[@id='Remarks']";
		public const string BtnSavePopUp = $"//div[@ng-form='frmAddDtl']//div[@ng-click='OK()']";
	}
}
