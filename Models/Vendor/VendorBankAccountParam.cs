namespace EPaymentVoucher.Models.Vendor
{
	public class VendorBankAccountParam : VendorBaseModels
	{
		public string BankAccountCountry { get; set; } = string.Empty;
		public string BankAccountAliasName { get; set; } = string.Empty;
		public string BankAccountBankName { get; set; } = string.Empty;
		public string BankAccountBranch { get; set; } = string.Empty;
		public string BankAccountNo { get; set; } = string.Empty;
		public string BankAccountName { get; set; } = string.Empty;
		public string BankAccountCurrency { get; set; } = string.Empty;
		public string BankAccountSwiftCode { get; set; } = string.Empty;
		public bool BankAccountIsDefault { get; set; } = false;
		public string BankAccountEvidencePath { get; set; } = string.Empty;
	}
}
