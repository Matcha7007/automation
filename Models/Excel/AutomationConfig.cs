using EPaymentVoucher.Models.BankConfirmation;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Models.Vendor;

namespace EPaymentVoucher.Models.Excel
{
	public class AutomationConfig
	{
		public virtual List<User> Users { get; set; } = [];
		public virtual List<TestPlan> TestPlans { get; set; } = [];
		public virtual List<VendorDatabase> VendorDatabases { get; set; } = [];
		public virtual List<VendorDatabaseTask> VendorDatabaseTasks { get; set; } = [];
		public virtual List<PVReimbursementModels> PVReimbursements { get; set; } = [];
		public virtual List<PVAdvanceModels> PVAdvances { get; set; } = [];
		public virtual List<PVSettlementModels> PVSettlements { get; set; } = [];
		public virtual List<PVVendorModels> PVVendors { get; set; } = [];
		public virtual List<PVTask> PVTasks { get; set; } = [];
		public virtual List<PVSalesForceModels> PVRDDs { get; set; } = [];
		public virtual List<PVSalesForceModels> PVSalesForces { get; set; } = [];
		public virtual List<PVAgencyCommissionModels> PVAgencyCommissions { get; set; } = [];
		public virtual List<PVBancassuranceBMIModels> PVBancassuranceBMIs { get; set; } = [];
		public virtual List<PVBancassuranceBTNModels> PVBancassuranceBTNs { get; set; } = [];
		public virtual List<PVPolicyHolderModels> PVPolicyHolders { get; set; } = [];
		public virtual List<PVInternalBankModels> PVInternalBanks { get; set; } = [];
		public virtual List<PVSpecialCustomerModels> PVSpecialCustomers { get; set; } = [];
		public virtual List<BankConfirmationModels> RejectBankConfirmations { get; set; } = [];
	}
}
