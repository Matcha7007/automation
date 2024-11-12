namespace EPaymentVoucher.Models.Vendor
{
	public class VendorDatabaseTask : ModelBase
	{
		public virtual VendorDatabaseSearchTask ParamSearch { get; set; } = new();
        public int Sequence { get; set; }
        public string Action { get; set; } = string.Empty;
		public string Actor { get; set; } = string.Empty;
		public string ReturnTo { get; set; } = string.Empty;
		public string Notes { get; set; } = string.Empty;
		public string NewVendorCode { get; set; } = string.Empty;
		public string Result {  get; set; } = string.Empty;
		public string Screenshot { get; set; } = string.Empty;
    }

	public class VendorDatabaseSearchTask
	{
		public string VendorCode { get; set; } = string.Empty;
		public string VendorName { get; set; } = string.Empty;
		public string BusinessName { get; set; } = string.Empty;
	}
}
