using EPaymentVoucher.Locators.PaymentVoucher;

namespace EPaymentVoucher.Locators.BankConfirmation
{
	public class BankConfirmationXPath
	{
		public const string UrlPendingPaymentList = $"{PVListXPath.UrlPaymentVoucher}/PendingPaymentList";
		public const string UrlBankConfirmationList = $"{PVListXPath.UrlPaymentVoucher}/BankConfirmationList";

		public const string RequestNo = "//input[@id='RequestNo']";
		public const string PVNo = "//input[@id='PVNo']";
		public const string PaymentType = "//select[contains(@ng-model,'PaymentType')]";
		public const string Purpose = "//select[contains(@ng-model,'Purpose')]";
		public const string Payee = "//input[@name='Payee']";
		public const string DateFrom = "//input[@id='datePickerAppDate1']";
		public const string DateTo = "//input[@id='datePickerAppDate2']";
		public const string BankProfile = "//select[contains(@ng-model,'BankProfileId')]";
		public const string Currency = "//select[contains(@ng-model,'CurrencyCode')]";
		public const string BtnSearch = "//div[@id='btnSearch']";
		public const string BtnClear = "//div[@ng-click='ClearData()']";

		// author untuk param search bank confirmation
		public const string Author = "//input[contains(@ng-model,'Author_Contain')]";

		public const string RequestPaymentVoucherList = "//div[@ui-grid='gridReqPVList']//div[contains(@ng-class, 'row.isSelected')]";
		public const string DetailTransactionList = "//div[@ui-grid='gridDetailTransList']//div[contains(@ng-class, 'row.isSelected')]";
		public const string DetailTransactionAll = "//div[@ui-grid='gridDetailTransList']//div[@ng-model='grid.selection.selectAll']";
		public const string BtnSelectDetailTransaction = "//div[@ng-click='SelectPVDetail()']";
		public const string TotalTransactionAll = "//div[@ui-grid='gridSelectTransactionList']//div[@ng-model='grid.selection.selectAll']";
		public const string BtnDeselectTotalTransaction = "//div[@ng-click='DeselectTransaction()']";
		public const string TransferDate = "//input[@id='transferDatePicker']";
		public const string BankSource = "//select[@id='bankSource']";
		public const string Rate = "//input[@id='rate']";
		public const string BtnCalculate = "//button[@ng-click='ReportingCurrencyCalculate()']";
		public const string BtnReturn = "//div[@ng-click='Reject()']";
	}
}
