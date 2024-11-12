namespace EPaymentVoucher.Models
{
    public class AppConfig
    {
        public string Url { get; set; } = string.Empty;
        public string LogPath { get; set; } = string.Empty;
        public string ExcelConfigPath { get; set; } = string.Empty;
        public string FileTestPath { get; set; } = string.Empty;
        public string ScreenCapturePath { get; set; } = string.Empty;
        public string TestData { get; set; } = string.Empty;
		public int WaitElementInSecond { get; set; }
		public int WaitWriteResultInSecond { get; set; }
	}

	public static class GlobalConfig
	{
		public static AppConfig Config { get; set; } = new();
	}
}
