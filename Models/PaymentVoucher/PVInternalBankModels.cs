namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVInternalBankModels : PVBaseModels
	{
		public string FromBankAccountNo { get; set; } = string.Empty;
		public string ToBankAccountNo { get; set; } = string.Empty;
		public string CostCenter { get; set; } = string.Empty;
		public string LongDesc { get; set; } = string.Empty;
		public string ShortDesc { get; set; } = string.Empty;
		public string Currency { get; set; } = string.Empty;
		public string Amount { get; set; } = string.Empty;
		public string DetailAttachmentPath { get; set; } = string.Empty;
		public string DetailRemarks { get; set; } = string.Empty;
	}
}
