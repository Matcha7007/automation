namespace EPaymentVoucher.Models.Vendor
{
	public class VendorDatabase : VendorBaseModels
	{
		public virtual VendorDatabaseSearch SearchVendor { get; set; } = new();
		public string VendorCode { get; set; } = string.Empty;
		public string VendorName { get; set; } = string.Empty;
		public string VendorType { get; set; } = string.Empty;
		public string Country { get; set; } = string.Empty;
		public string City { get; set; } = string.Empty;
		public string Address { get; set; } = string.Empty;
		public string Phone1 { get; set; } = string.Empty;
		public string Phone2 { get; set; } = string.Empty;
		public string Fax { get; set; } = string.Empty;
		public string Email { get; set; } = string.Empty;
		public bool IsHaveCertificateOfDomicile { get; set; } = false;
		public string CoDAttachmentPath { get; set; } = string.Empty;
		public string CoDExpiryDate { get; set; } = string.Empty;
		public bool IsTaxableEntrepreneur { get; set; } = false;
		public string SPPKPPath { get; set; } = string.Empty;
		public bool IsUnderWithholdingTax { get; set; } = false;
		public string Remarks { get; set; } = string.Empty;
		public string AttachmentPath { get; set; } = string.Empty;
		public string AttachmentDescription { get; set; } = string.Empty;

		// Vendor Tax
		public bool IsHaveVendorTaxId { get; set; } = false;
		public string TaxIdNo { get; set; } = string.Empty;
		public string TaxIdName { get; set; } = string.Empty;
		public string TaxIdAddress { get; set; } = string.Empty;
		public string TaxIdAttachmentPath { get; set; } = string.Empty;
		public string TaxIdRemarks { get; set; } = string.Empty;
		public virtual List<VendorContactPersonParam> ContactPersons { get; set; } = new();
		public virtual List<VendorBankAccountParam> BankAccounts { get; set; } = new();
		public virtual List<VendorBusinessNameParam> BusinessNames { get; set; } = new();
    }
}
