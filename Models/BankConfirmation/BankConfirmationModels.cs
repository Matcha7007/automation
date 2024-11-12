namespace EPaymentVoucher.Models.BankConfirmation
{
	public class BankConfirmationBase : ModelBase
	{
		public string RequestNo { get; set; } = string.Empty;
		public string PVNo { get; set; } = string.Empty;
        public string PaymentType { get; set; } = string.Empty;
		public string Purpose { get; set; } = string.Empty;
        public string Payee { get; set; } = string.Empty;
        public string BankProfile { get; set; } = string.Empty;
        public string Currency { get; set; } = string.Empty;
	}

	public class PendingPaymentModels : BankConfirmationBase
	{
		public string ApproveDateFrom { get; set; } = string.Empty;
		public string ApproveDateTo { get; set; } = string.Empty;
		public string TransferType { get; set; } = string.Empty;
		public string BankSource { get; set; } = string.Empty;
    }

	public class BankConfirmationModels : BankConfirmationBase
	{
		public string Author { get; set; } = string.Empty;
		public string PaymentDateFrom { get; set; } = string.Empty;
		public string PaymentDateTo { get; set; } = string.Empty;
		public string TransferDate { get; set; } = string.Empty;
		public string BankSource { get; set; } = string.Empty;
		public string Rate { get; set; } = string.Empty;
	}

	public class TransferTypeConst
	{
		public const string Mass = "Mass";
		public const string Single = "Single";
	}
}
