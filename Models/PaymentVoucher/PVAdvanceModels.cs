namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVAdvanceModels : PVBaseModels
	{
		public string RequestorName { get; set; } = string.Empty;
		public string BankName { get; set; } = string.Empty;
		public string RequestorBankAccountName { get; set; } = string.Empty;
		public string RequestorBankAccountNo { get; set; } = string.Empty;
		public string StartDate { get; set; } = string.Empty;
		public string EndDate { get; set; } = string.Empty;
        public string Amount { get; set; } = string.Empty;
        public string Reason { get; set; } = string.Empty;
	}
}
