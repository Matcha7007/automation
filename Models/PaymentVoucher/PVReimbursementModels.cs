namespace EPaymentVoucher.Models.PaymentVoucher
{
    public class PVReimbursementModels : PVBaseModels
    {
        public string RequestorName { get; set; } = string.Empty;
        public string BankName { get; set; } = string.Empty;
        public string RequestorBankAccountName { get; set; } = string.Empty;
        public string RequestorBankAccountNo { get; set; } = string.Empty;
        public virtual List<PVDetailTransactionModels> AddDetails { get; set; } = new();
    }
}
