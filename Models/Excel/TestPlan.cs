namespace EPaymentVoucher.Models.Excel
{
	public class TestPlan : ModelBase
	{
        public string ModuleName { get; set; } = string.Empty;
        public string TestCase { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string DateTest { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string ScreenCapture { get; set; } = string.Empty;
        public string TestData { get; set; } = string.Empty;
        public string UserLogin { get; set; } = string.Empty;
		public int FromTestCaseId { get; set; }
		public string RequestResult { get; set; } = string.Empty;
    }
}
