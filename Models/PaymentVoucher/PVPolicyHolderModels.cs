namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVPolicyHolderModels : PVBaseModels
	{
		public string RequestorName { get; set; } = string.Empty;
		public string Payee { get; set; } = string.Empty;
		public string PolicyNo { get; set; } = string.Empty;
		public string SPAJNo { get; set; } = string.Empty;
		public string BussinessName { get; set; } = string.Empty;
		public string LongDesc { get; set; } = string.Empty;
		public string ShortDesc { get; set; } = string.Empty;
		public string Product { get; set; } = string.Empty;
		public string Distribution { get; set; } = string.Empty;
		public string Country { get; set; } = string.Empty;
		public string BankName { get; set; } = string.Empty;
		public string BankAccountNo { get; set; } = string.Empty;
		public string BankAccountName { get; set; } = string.Empty;
		public string Branch { get; set; } = string.Empty;
		public string Currency { get; set; } = string.Empty;
		public string GrossAmount { get; set; } = string.Empty;
		public string TaxAmount { get; set; } = string.Empty;
		public string ChargesAdmin { get; set; } = string.Empty;
		public string ChargesMedical { get; set; } = string.Empty;
		public string ForPremiumPayment { get; set; } = string.Empty;
		public string DetailAttachmentPath { get; set; } = string.Empty;
		public string TaxIdNo { get; set; } = string.Empty;
		public string TaxIdName { get; set; } = string.Empty;
		public string TaxIdAddress { get; set; } = string.Empty;
		public string TaxIdAttachmentPath { get; set; } = string.Empty;
		public string DetailRemarks { get; set; } = string.Empty;
	}
}
