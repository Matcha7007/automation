namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVSalesForceXPath
	{
		public const string CostCenterName = "//select[@id='CostCenterName']";
		public const string Payment = "//select[@id='Payment']";
		public const string ShortDescription = "//input[@id='ShortDescription']";
		public const string RowFile = "//div[@class='row'][3]";
		public const string BtnFile = $"{RowFile}{PVHeaderXPath.BaseBtnFile}";
		public const string BtnUploadFile = $"{RowFile}{PVHeaderXPath.BaseBtnUploadFile}";
		public const string IsProposalTrue = "//input[@name='IsProposal'][@ng-value='true']";
		public const string IsProposalFalse = "//input[@name='IsProposal'][@ng-value='false']";
		public const string RowProposal = "//div[@class='row'][5]";
		public const string BtnProposal = $"{RowProposal}{PVHeaderXPath.BaseBtnFile}";
		public const string PeriodOfPayment = "//input[@id='PeriodofPayment']";
		public const string BtnBusinessName = "//div[@id=' vendor-business-name']";
		public const string BeneficiaryCustomerType = "//select[@id='BeneficiaryCustomerType']";
		public const string TotalGross = "//input[@id='TotalGrossAmount']";
		public const string TotalNett = "//input[@id='TotalNetAfterTax']";
	}
}
