namespace EPaymentVoucher.Models.Vendor
{
	public class VendorBaseModels : ModelBase
	{
		public int Sequence { get; set; }
		public string DataFor { get; set; } = string.Empty;
	}
}
