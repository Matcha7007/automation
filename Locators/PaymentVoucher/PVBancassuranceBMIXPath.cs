namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVBancassuranceBMIXPath
	{
		public const string BaseLocator = "//div[@ng-form='formContainer']/div[3]";
		public const string BtnPaymentRelated = $"{BaseLocator}/div[2]{PVHeaderXPath.BaseBtnFile}";
		public const string BtnGLRelated = $"{BaseLocator}/div[4]{PVHeaderXPath.BaseBtnFile}";
		public const string BtnUpload = PVHeaderXPath.BaseBtnUploadFile;
		public const string CostCenter = $"//select[@ng-model='selectedCostCenter']";
	}

	public class BankStafBMIXPath
	{
		public const string BaseLocator = PVBancassuranceBMIXPath.BaseLocator;
		public const string BtnPaymentRelated = PVBancassuranceBMIXPath.BtnPaymentRelated;
		public const string BtnGLRelated = $"{BaseLocator}/div[4]/div[1]{PVHeaderXPath.BaseBtnFile}";
		public const string BtnDetailBMI = $"{BaseLocator}/div[4]/div[2]{PVHeaderXPath.BaseBtnFile}";
		public const string BtnUpload = PVHeaderXPath.BaseBtnUploadFile;
		public const string CostCenter = PVBancassuranceBMIXPath.CostCenter;
	}
}
