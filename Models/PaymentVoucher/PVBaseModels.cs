namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVBaseModels : ModelBase
	{
        public int Sequence { get; set; }
        public string DataFor { get; set; } = string.Empty;
		public string Title { get; set; } = string.Empty;
		public string PaymentType { get; set; } = string.Empty;
		public string PaymentPurpose { get; set; } = string.Empty;
		public string AttachmentPath { get; set; } = string.Empty;
		public string AttachmentDescription { get; set; } = string.Empty;
		public string Remarks { get; set; } = string.Empty;
	}
}
