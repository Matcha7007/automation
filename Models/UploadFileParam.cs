namespace EPaymentVoucher.Models
{
	public class UploadFileParam
	{
		public string FilePath { get; set; } = null!;
		public string XPathBtnUpload { get; set; } = null!;
        public string DialogTitle { get; set; } = null!;
		public string ModuleName { get; set; } = null!;
        public string? Section { get; set; }
    }
}
