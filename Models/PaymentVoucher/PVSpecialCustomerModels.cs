namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVSpecialCustomerModels : PVBaseModels
	{
		public string RequestorName { get; set; } = string.Empty;
		public string Payee { get; set; } = string.Empty;
		public string PayeeType { get; set; } = string.Empty;
		public string Country { get; set; } = string.Empty;
		public bool IsTaxable { get; set; }
		public string SKPPPath { get; set; } = string.Empty;
		public bool IsCoD { get; set; }
		public string CoDPath { get; set; } = string.Empty;
		public string CoDExpiryDate { get; set; } = string.Empty;
		public bool IsUWTax { get; set; }
		public string TaxIdNo { get; set; } = string.Empty;
		public string TaxIdName { get; set; } = string.Empty;
		public string TaxIdAddress { get; set; } = string.Empty;
		public string TaxIdAttachmentPath { get; set; } = string.Empty;
		public string BankName { get; set; } = string.Empty;
		public string BankAccountNo { get; set; } = string.Empty;
		public string BankAccountName { get; set; } = string.Empty;
		public string Branch { get; set; } = string.Empty;
		public virtual PVSpecialCustomerDetailModels Detail { get; set; } = new();
	}

	public class PVSpecialCustomerDetailModels
	{
		public string BussinessName { get; set; } = string.Empty;
		public string CostCenter { get; set; } = string.Empty;
		public string LongDesc { get; set; } = string.Empty;
		public string ShortDesc { get; set; } = string.Empty;
		public string GrossAmount { get; set; } = string.Empty;
		public string Currency { get; set; } = string.Empty;
		public bool IsTaxExemption { get; set; }
		public string SKBPath { get; set; } = string.Empty;
		public string AttachmentPath { get; set; } = string.Empty;
		public string Remarks { get; set; } = string.Empty;
	}
}
