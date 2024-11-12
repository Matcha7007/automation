namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVBancassuranceBTNXPath
	{
		public const string BaseBtnFile = PVHeaderXPath.BaseBtnFile;
		public const string BaseBtnUploadFile = PVHeaderXPath.BaseBtnUploadFile;

		public const string BaseLocator = "//div[@ng-form='formContainer']/div[3]";
		public const string BtnPaymentRelated = $"{BaseLocator}/div[2]{BaseBtnFile}";
		public const string BtnGLRelated = $"{BaseLocator}/div[4]{BaseBtnFile}";
		public const string CostCenter = $"//select[@ng-model='selectedCostCenter']";
	}
}
