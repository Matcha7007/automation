namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVSalesForceModels : PVBaseModels
	{
		public string CostCenterName { get; set; } = string.Empty;
		public string BusinessName { get; set; } = string.Empty;
		public string Payment { get; set; } = string.Empty;
		public string ShortDescription { get; set; } = string.Empty;
		public string FilePath { get; set; } = string.Empty;
		public bool IsProposal { get; set; }
		public string ProposalPath { get; set; } = string.Empty;
		public string BeneficiaryCustomerType { get; set; } = string.Empty;
		public string PeriodOfPayment { get; set; } = string.Empty;
	}
}
