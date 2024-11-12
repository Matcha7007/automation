namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVSettlementModels : PVBaseModels
	{
		public string RequestorName { get; set; } = string.Empty;
		public string PaymentVoucherNo { get; set; } = string.Empty;
		public virtual List<PVDetailTransactionModels> AddDetails { get; set; } = new();
	}
}
