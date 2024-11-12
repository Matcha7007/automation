namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVTask : PVBaseModels
	{
		public string Action { get; set; } = string.Empty;
		public string Actor { get; set; } = string.Empty;
		public string RequestNo { get; set; } = string.Empty;
        public string RequestDateFrom { get; set; } = string.Empty;
        public string RequestDateTo { get; set; } = string.Empty;
        public string Payee { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public bool IsPaymentViaH2H { get; set; }
        public string ReturnTo { get; set; } = string.Empty;
        public string Notes { get; set; } = string.Empty;
		public bool IsPayment { get; set; } = false;
		public string Result { get; set; } = string.Empty;
		public string Screenshot { get; set; } = string.Empty;
	}
}
