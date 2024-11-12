namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVVendorModels : PVBaseModels
	{
        public string Payee { get; set; } = string.Empty;
        public bool IsRecurring { get; set; }
		public string PO { get; set; } = string.Empty;
		public string GR { get; set; } = string.Empty;
		public virtual List<PVDetailTransactionLongModels> AddDetails { get; set; } = new();
	}
}
