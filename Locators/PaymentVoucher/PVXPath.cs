namespace EPaymentVoucher.Locators.PaymentVoucher
{
	public class PVPopUpEmployeeXPath
	{
		public const string ListCellRequestor = "//div[@role='gridcell'][1]";
		public const string Requestor = "//div[@class='modal-body']//input[contains(@ng-model, 'Requestor')]";
		public const string Department = "//div[@class='modal-body']//input[contains(@ng-model, 'Department')]";
		public const string Position = "//div[@class='modal-body']//input[contains(@ng-model, 'Position')]";
		public const string BtnSearch = "//div[@id='btn-search']";
	}

	public class PVPopUpBusinessNameXPath
	{
		public const string ListCellCategory = "//div[@role='gridcell'][1]";
		public const string Category = "//div[contains(@class, 'modal-body')]//input[@id='Category']";
		public const string Name = "//div[contains(@class, 'modal-body')]//input[@id='Name']";
		public const string Type = "//div[contains(@class, 'modal-body')]//input[@id='Type']";
		public const string BtnSearch = "//div[@id='btn-search']";
		public const string BtnClear = "//div[@id='businessNameCancel']";
	}

	public class PVPopUpPayeeXPath
	{
		public const string ListCellVendorCode = "//div[@role='gridcell'][1]";
		public const string VendorCode = "//div[contains(@class, 'modal-body')]//input[contains(@ng-model, 'VendorCode')]";
		public const string VendorName = "//div[contains(@class, 'modal-body')]//input[contains(@ng-model, 'VendorName')]";
		public const string BtnSearch = "//div[@id='btn-search']";
	}

	public class PVPopUpDetailTransactionXPath
	{
		public const string BtnAddListBusinessName = "//a[@ng-click='open()']//div[@id='vendor-business-name']";
		public const string ListCellPopup = "//div[@role='gridcell'][2]";
		public const string CostCenter = "//select[@name='CostCenter']";
		public const string AccrualCode = "//select[@name='Accrual']";
		public const string LongDescription = "//textarea[@id='LongDescription']";
		public const string ShortDescription = "//input[@name='ShortDescription']";
		public const string Date = "//input[@id='Date']";
		public const string Amount = "//input[@id='Amount']";
		public const string Attachment = "//input[@id='AttachmentObject']";
		public const string RowAttachment = "//div[contains(@class, 'modal-body')]/div[9]";
		public const string BtnAttachment = $"{RowAttachment}//span[contains(@class, 'custom-upload')]";
		public const string IsProposalTrue = "//input[@name='IsProposal'][@value='true']";
		public const string IsProposalFalse = "//input[@name='IsProposal'][@value='false']";
		public const string RowProposal = "//div[contains(@class, 'modal-body')]/div[11]";
		public const string BtnProposal = $"{RowProposal}//span[contains(@class, 'custom-upload')]";
		public const string Proposal = "//input[@id='ProposalObject']";
		public const string BtnSave = "//div[@id='submit-vendor-edit']";
	}

	public class PVPopUpDetailTransactionLongXPath
	{
		public const string BtnRequestorName = "//dirpopupwindow[contains(@url, 'RequestorInformation')]//a[@ng-click='open()']";
		public const string CostCenterName = "//select[@ng-model='SelectedCostCenter']";
		public const string AccrualCode = "//select[@ng-model='SelectedAccrual']";
		public const string BtnBusinessName = "//dirpopupwindow[contains(@url, 'PVBusinessName')]//a[@ng-click='open()']";
		public const string ShortDescription = "//input[@id='ShortDescription']";
		public const string LongDescription = "//textarea[@id='LongDescription']";
		public const string IsProposalTrue = "//input[@name='IsProposal'][@ng-value='true']";
		public const string IsProposalFalse = "//input[@name='IsProposal'][@ng-value='false']";
		public const string BtnProposal = $"//div[contains(@ng-show, 'IsProposal')]{PVHeaderXPath.BaseBtnFile}";
		public const string InvoiceNo = "//input[@id='InvoiceNo']";
		public const string InvoiceDate = "//input[@id='InvoiceDate']";
		public const string InvoiceDueDate = "//input[@id='InvoiceDueDate']";
		public const string BtnInvoiceAttachment = $"//div[contains(@ng-class, 'InvoiceAttachment')]{PVHeaderXPath.BaseBtnFile}";
		public const string BankAccountNo = "//select[@ng-model='selectedBankOption']";
		public const string GrossAmount = "//input[@id='GrossAmount']";
		public const string Remarks = "//textarea[@id='Remarks']";
		public const string BtnSavePopUp = $"//div[@ng-click='OK(popUpform)']";
	}

	public class PVPopUpDetailDetailTransactionXPath
	{
		public const string BtnAddDetails = "//div[contains(@ng-show,'isNew')]/dirpopupwindow[@initial-value='NewItem()']//a[@ng-click='open()']";
		public const string Date = "//input[@id='Date']";
		public const string Description = "//select[@ng-model='SelectedDescription']";
		public const string Amount = "//input[@id='Amount']";
		public const string Currency = "//select[@id='Currency']";
		public const string Rate = "//input[@id='Rate']";
		public const string BtnAttachment = "//div[contains(@ng-class, 'SubPopUpform.Attachment')]//span[contains(@class, 'custom-upload')]";
		public const string BtnSave = "//div[@ng-click='OK(SubPopUpform)']";
		public const string AttributeClassDate = "form-control has-feedback hasDatepicker ng-touched ng-not-empty ng-dirty ng-valid-parse ng-valid ng-valid-required";
	}

	public class PVOtherDetailXPath
	{
		public const string BtnAttachment = "//div[@id='other-detail-upload-attachment']//span[contains(@class, 'custom-upload')]";
		public const string Description = "//textarea[contains(@ng-model, 'Description')]";
		public const string BtnAddAttachment = "//div[@ng-click='AddAttachment()']";
		public const string Remarks = "//textarea[contains(@ng-model, 'Remarks')]";
		public const string BtnSubmit = "//div[@ng-click='SubmitPVForm()']";
		public const string ListRowDeleteAttachment = "//tr[contains(@ng-repeat, 'Attachments')]//a[@ng-click='DeleteItem($index)']";
	}

	public class PVApproval
	{
		public const string PaymentViaTrue = "//input[contains(@name, 'paymentVia')][@value='true'][not(@disabled)]";
		public const string PaymentViaFalse = "//input[contains(@name, 'paymentVia')][@value='false'][not(@disabled)]";
		public const string Action = "//select[@ng-model='selectedAction']";
		public const string ReturnTo = "//select[@ng-model='SelectedReturnTo']";
		public const string Notes = "//textarea[@id='Notes']";
		public const string BtnSubmit = "//div[@ng-click='SubmitPVForm(mainFormContainer)']";
	}

	public class PVHeaderXPath
	{
		public const string Title = PVListXPath.Title;
		public const string PaymentType = PVListXPath.HeaderPaymentType;
		public const string PaymentPurpose = PVListXPath.HeaderPaymentPurpose;
		public const string BtnNext = PVListXPath.HeaderBtnNext;
		public const string BtnCancel = PVListXPath.HeaderBtnCancel;

		public const string BaseBtnFile = $"//span[contains(@class, 'custom-upload')]";
		public const string BaseBtnUploadFile = $"//div[@ng-click='Upload()']";
	}
}
