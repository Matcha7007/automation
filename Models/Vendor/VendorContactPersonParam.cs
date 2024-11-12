namespace EPaymentVoucher.Models.Vendor
{
	public class VendorContactPersonParam : VendorBaseModels
	{
		public string ContactPersonName { get; set; } = string.Empty;
		public string ContactPersonPhone1 { get; set; } = string.Empty;
		public string ContactPersonPhone2 { get; set; } = string.Empty;
		public string ContactPersonEmail { get; set; } = string.Empty;
	}
}
