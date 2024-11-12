namespace EPaymentVoucher.Locators.Vendor
{
	public partial class VendorListXPath
	{
		public const string TabVendorList = "//a[@href='/Vendor/VendorList']";
		public const string VendorCode = "//input[@id='VendorCode']";
		public const string VendorName = "//input[@id='VendorName']";
		public const string BusinessName = "//input[@id='BusinessName']";
		public const string BtnSearch = "//a[@id='searchBtn']";
		public const string BtnClear = "//div[@id='clearBtn']";
		public const string BtnCreate = "//a[@href='/Vendor/Create']";
		public const string BtnEdit = "//tr//a[contains(@href, '/Vendor/Edit')]";

		#region Form
		public const string FormVendor = "//form[@id='form-vendor-database']";

		#region Profile
		public const string TabVendorProfile = $"{FormVendor}//div[@id='vendorProfile']";

		public const string FormVendorName = $"{TabVendorProfile}//input[@id='Profile_Name']";
		public const string FormVendorType = $"{TabVendorProfile}//select[@id='Profile_Type']";
		public const string FormVendorTypeOpt = $"{FormVendorType}/option";
		public const string FormCountry = $"{TabVendorProfile}//select[@id='Profile_Country']";
		public const string FormCountryOpt = $"{FormCountry}/option";
		public const string FormCoDTrue = $"{TabVendorProfile}//input[@id='CoDtrue']";
		public const string FormCoDFalse = $"{TabVendorProfile}//input[@id='CoDfalse']";
		public const string FormBtnCoDAttachment = $"{TabVendorProfile}//div[@id='divCoDFileUpload']";
		public const string FormCoDAttachment = $"{TabVendorProfile}//input[@id='Profile_CoDAttachment']";
		public const string FormCoDExpDate = $"{TabVendorProfile}//input[@id='Profile_CoDExpiry']";
		public const string FormTaxableTrue = $"{TabVendorProfile}//input[@id='TaxableTrue']";
		public const string FormTaxableFalse = $"{TabVendorProfile}//input[@id='TaxableFalse']";
		public const string FormBtnSPPKPAttachment = $"{TabVendorProfile}//div[@id='taxableEntrepeneurFileUpload']";
		public const string FormSPPKPAttachment = $"{TabVendorProfile}//input[@id='Profile_SPPKP']";
		public const string FormUWTTrue = $"{TabVendorProfile}//input[@id='Profile_WithholdingTax'][@value='true']";
		public const string FormUWTFalse = $"{TabVendorProfile}//input[@id='Profile_WithholdingTax'][@value='false']";
		public const string FormRemarks = $"{TabVendorProfile}//textarea[@id='Profile_Remarks']";
		public const string FormAttachment = $"{TabVendorProfile}//div[@id='attachment-partial']";
		public const string FormAttachmentBtnSelectFile = $"{FormAttachment}/div[1]//div[contains(@class, 'custom-upload')]";
		public const string FormAttachmentSelectFile = $"{FormAttachment}//input[@id='Attachments_Attachment_VendorId']";
		public const string FormAttachmentDescription = $"{FormAttachment}//textarea[@id='Attachments_Attachment_Description']";
		public const string FormAttachmentBtnUpload = $"{FormAttachment}//div[@id='upload-attachment']";
		public const string FormCity = $"{TabVendorProfile}//input[@id='Profile_City']";
		public const string FormAddress = $"{TabVendorProfile}//textarea[@id='Profile_Address']";
		public const string FormPhone1_Front = $"{TabVendorProfile}//input[@data-fill='Profile_Phone1'][@data-id='front-phone']";
		public const string FormPhone1_Main = $"{TabVendorProfile}//input[@data-fill='Profile_Phone1'][@data-id='main-phone']";
		public const string FormPhone1_Back = $"{TabVendorProfile}//input[@data-fill='Profile_Phone1'][@data-id='back-phone']";
		public const string FormPhone2_Front = $"{TabVendorProfile}//input[@data-fill='Profile_Phone2'][@data-id='front-phone']";
		public const string FormPhone2_Main = $"{TabVendorProfile}//input[@data-fill='Profile_Phone2'][@data-id='main-phone']";
		public const string FormPhone2_Back = $"{TabVendorProfile}//input[@data-fill='Profile_Phone2'][@data-id='main-phone']";
		public const string FormFaxNo_Front = $"{TabVendorProfile}//input[@data-fill='FaxNo'][@data-id='front-fax']";
		public const string FormFaxNo_Main = $"{TabVendorProfile}//input[@data-fill='FaxNo'][@data-id='main-fax']";
		public const string FormEmail = $"{TabVendorProfile}//input[@id='Profile_Email']";

		//TAX ID
		public const string FormTaxIdTrue = $"{TabVendorProfile}//input[@id='TaxIdtrue']";
		public const string FormTaxIdFalse = $"{TabVendorProfile}//input[@id='TaxIdfalse']";
		public const string FormTaxIdNo = $"{TabVendorProfile}//input[@id='Profile_TaxNo']";
		public const string FormTaxIdName = $"{TabVendorProfile}//input[@id='Profile_TaxName']";
		public const string FormTaxIdAddress = $"{TabVendorProfile}//textarea[@id='Profile_TaxAdress']";
		public const string FormTaxIdBtnAttachment = $"{TabVendorProfile}//div[@id='divTaxIdFileUpload']";
		public const string FormTaxIdAttachment = $"{TabVendorProfile}//input[@id='Profile_TaxAttachment']";
		public const string FormTaxIdRemarks = $"{TabVendorProfile}//textarea[@id='Profile_TaxRemarks']";

		public const string FormBtnSaveAsDraft = $"{TabVendorProfile}//div[@data-name='save-draft']";
		public const string FormBtnCancel = $"{TabVendorProfile}//a[@href='/Vendor/VendorList']";
		public const string FormBtnNext = $"{TabVendorProfile}//div[@id='next-from-profile']";
		#endregion

		#region Detail
		public const string TabVendorDetail = $"{FormVendor}//div[@id='vendorDetail']";
		public const string FormModal = "//div[@id='master-modal']";
		public const string FormModalSearch = "//div[@id='search-modal']";

		#region Contact Person
		public const string FormDBtnAddContactPerson = $"{TabVendorDetail}//div[@id='add-contact-person']";
		public const string FormDBtnDeleteContactPerson = $"{TabVendorDetail}//tr//a[contains(@onclick, 'contactpopDelete')]";
		public const string FormDCPName = $"{FormModal}//input[@id='Name']";
		public const string FormDCPPhone1_Front = $"{FormModal}//input[@data-fill='Phone1'][@data-id='front-phone']";
		public const string FormDCPPhone1_Main = $"{FormModal}//input[@data-fill='Phone1'][@data-id='main-phone']";
		public const string FormDCPPhone1_Back = $"{FormModal}//input[@data-fill='Phone1'][@data-id='back-phone']";
		public const string FormDCPPhone2_Front = $"{FormModal}//input[@data-fill='Phone2'][@data-id='front-phone']";
		public const string FormDCPPhone2_Main = $"{FormModal}//input[@data-fill='Phone2'][@data-id='main-phone']";
		public const string FormDCPPhone2_Back = $"{FormModal}//input[@data-fill='Phone2'][@data-id='back-phone']";
		public const string FormDCPEmail = $"{FormModal}//input[@id='Email']";
		public const string FormDCPBtnAdd = $"{FormModal}//div[@id='submit-cp']";
		public const string FormDCPBtnCancel = $"{FormModal}//input[@id='contactCancell']";
		#endregion

		#region Bank Account
		public const string FormDBtnAddBankAccount = $"{TabVendorDetail}//div[@id='add-bank-account']";
		public const string FormDBtnDeleteBankAccount = $"{TabVendorDetail}//tr//a[contains(@onclick, 'RealBankAccount')]";
		public const string FormDBACountry = $"{FormModal}//select[@id='BankCountry']";
		public const string FormDBACountryOpt = $"{FormDBACountry}/option";
		public const string FormDBAAliasName = $"{FormModal}//input[@id='AliasName']";
		public const string FormDBABankName = $"{FormModal}//select[@id='BankIndo']";
		public const string FormDBABankNameInput = $"{FormModal}//input[@id='BankName']";
		public const string FormDBABranch = $"{FormModal}//input[@id='Branch']";
		public const string FormDBABankNameOpt = $"{FormDBABankName}/option";
		public const string FormDBAAccountNo = $"{FormModal}//input[@id='AccountNo']";
		public const string FormDBAAccountName = $"{FormModal}//input[@id='AccountName']";
		public const string FormDBACurrency = $"{FormModal}//select[@id='Currency']";
		public const string FormDBACurrencyOpt = $"{FormDBACurrency}/option";
		public const string FormDBASwiftCode = $"{FormModal}//input[@id='SwiftCode']";
		public const string FormDBADefault = $"{FormModal}//label[@class='checkbox']/input[@id='Default'][@value='true']";
		public const string FormDBABtnUpload = $"{FormModal}//div[@class='upload']/div";
		public const string FormDBAUpload = $"{FormModal}//input[@id='Evidence']";
		public const string FormDBABtnAdd = $"{FormModal}//div[@id='submit-bank']";
		public const string FormDBABtnCancel = $"{FormModal}//div[@id='bankAccountCancell']";
		#endregion

		#region Vendor Business
		public const string FormDBtnAddVendorBusiness = $"//div[@id='add-vendor-BNSearch']";
		public const string FormDBtnDeleteVendorBusiness = $"{TabVendorDetail}//tr//a[contains(@onclick, 'RealBusiness')]";
		public const string FormDVBCategory = $"{FormModalSearch}//input[@id='Category']";
		public const string FormDVBName = $"{FormModalSearch}//input[@id='Name']";
		public const string FormDVBType = $"{FormModalSearch}//input[@id='TypeTextBox']";
		public const string FormDVBBtnSearch = $"{FormModalSearch}//div[@id='btn-search']";
		public const string FormDVBBtnCancel = $"{FormModalSearch}//div[@id='businessNameCancel']";
		public const string FormDVBTable = $"{FormModalSearch}//table[@id='businessnametable']";
		public const string FormDVBTableTr = $"{FormDVBTable}//tr";
		public const string FormDVBTableTrCategory = $"{FormDVBTableTr}/td[1]";
		public const string FormDVBTableTrBusinessName = $"{FormDVBTableTr}/td[2]";
		public const string FormDVBTableTrType = $"{FormDVBTableTr}/td[3]";
		public const string FormDVBTableTrChecked = $"{FormDVBTableTr}/td[7]/input";
		public const string FormDVBBtnSelect = $"{FormModalSearch}//div[@id='submit-vendorBusinessNameSelect']";
		public const string FormDVBNotes = $"{TabVendorDetail}//textarea[@id='SubmissionNote']";
		public const string mBussinessId = "//input[@name='Businesses[1].mVendorBusinessId']"; 
        #endregion

        public const string FormDetailBtnBack = $"{TabVendorDetail}//div[@data-target='#vendorProfile']";
		public const string FormDetailBtnSaveDraft = $"{TabVendorDetail}//div[@data-name='save-draft']";
		public const string FormDetailBtnSave = $"{TabVendorDetail}//div[@id='submitting']";
		public const string FormDetailBtnCancel = $"{TabVendorDetail}//a[@href='/Vendor/VendorList']";
		public const string FormDetailBtnNext = $"//div[@data-target='#approvalAction']";
		#endregion
		#endregion
	}
}
