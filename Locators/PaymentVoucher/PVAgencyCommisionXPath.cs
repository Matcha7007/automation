namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class ACSummaryXPath
	{
		public const string CostCenter = $"//select[@ng-model='selectedCostCenter']";
		public const string BaseBtnFile = PVHeaderXPath.BaseBtnFile;
		public const string BaseBtnUploadFile = PVHeaderXPath.BaseBtnUploadFile;
	}
	public class ACGAAllowancePTXPath
	{
		public const string BtnFile = ACSummaryXPath.BaseBtnFile;
		public const string BtnUploadFile = ACSummaryXPath.BaseBtnUploadFile;
	}
	public class ACFinancingXPath
	{
		public const string BtnFile = ACSummaryXPath.BaseBtnFile;
		public const string BtnUploadFile = ACSummaryXPath.BaseBtnUploadFile;
	}
	public class ACMGIXPath
	{
		public const string BtnFile = ACSummaryXPath.BaseBtnFile;
		public const string BtnUploadFile = ACSummaryXPath.BaseBtnUploadFile;
	}
	public class ACCommisionMEMonthXPath
	{
		public const string BaseLocator = "//div[@ng-form='formContainer']/div[3]";
		public const string BtnTaxCalculation = $"{BaseLocator}/div[2]{ACSummaryXPath.BaseBtnFile}";
		public const string BaseGLRelated = $"{BaseLocator}/div[4]";
		public const string BtnFirstYearCommission = $"{BaseGLRelated}/div[1]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnRenewalYearCommission = $"{BaseGLRelated}/div[2]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnReportRecOverride = $"{BaseGLRelated}/div[3]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnADGenDetail = $"{BaseGLRelated}/div[4]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnProducerBonusReport = $"{BaseGLRelated}/div[5]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnActivationAgencyLeaderBonusReport = $"{BaseGLRelated}/div[6]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnSummaryOverride = $"{BaseGLRelated}/div[7]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnGAAllowancePersonal = $"{BaseGLRelated}/div[8]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnSummaryAgentPaid = $"{BaseGLRelated}/div[9]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnCFGManualPayment = $"{BaseGLRelated}/div[10]{ACSummaryXPath.BaseBtnFile}";
		public const string BtnUpload = ACSummaryXPath.BaseBtnUploadFile;
	}
}
