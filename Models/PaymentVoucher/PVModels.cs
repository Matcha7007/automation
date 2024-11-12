namespace EPaymentVoucher.Models.PaymentVoucher
{
	public class PVDetailTransactionLongModels
	{
		public string RequestorName { get; set; } = string.Empty;
		public string CostCenter { get; set; } = string.Empty;
		public string AccrualCode { get; set; } = string.Empty;
		public string BusinessName { get; set; } = string.Empty;
		public string LongDescription { get; set; } = string.Empty;
		public string ShortDescription { get; set; } = string.Empty;
		public bool IsProposal { get; set; }
		public string ProposalPath { get; set; } = string.Empty;
		public string InvoiceNo { get; set; } = string.Empty;
		public string InvoiceDate { get; set; } = string.Empty;
		public string InvoiceDueDate { get; set; } = string.Empty;
		public string InvoiceAttachmentPath { get; set; } = string.Empty;
		public string BankAccountNo { get; set; } = string.Empty;
		public string GrossAmount { get; set; } = string.Empty;
		public string Remarks { get; set; } = string.Empty;
	}

	public class PVDetailTransactionModels
	{
		public string BusinessName { get; set; } = string.Empty;
		public string CostCenter { get; set; } = string.Empty;
		public string AccrualCode { get; set; } = string.Empty;
		public string LongDescription { get; set; } = string.Empty;
		public string ShortDescription { get; set; } = string.Empty;
		public string Date { get; set; } = string.Empty;
        public string Amount { get; set; } = string.Empty;
		public string AttachmentPath { get; set; } = string.Empty;
		public bool IsProposal { get; set; }
		public string ProposalPath { get; set; } = string.Empty;
		public virtual List<PVDetailDetailTransactionModels> Details { get; set; } = new();
    }

	public class PVDetailDetailTransactionModels
	{
		public string Date { get; set; } = string.Empty;
		public string Description { get; set; } = string.Empty;
		public string Amount { get; set; } = string.Empty;
		public string Currency { get; set; } = string.Empty;
		public string Rate { get; set; } = string.Empty;
		public string AttachmentPath { get; set; } = string.Empty;
	}
}
