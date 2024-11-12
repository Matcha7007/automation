using EPaymentVoucher.Models;
using EPaymentVoucher.Models.BankConfirmation;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Models.Vendor;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing.Imaging;
using System.Drawing;

using System.Runtime.InteropServices;
using System.Text;

namespace EPaymentVoucher.Utilities
{
	public static class JsonHelper
	{
		public static T DeserializeObjectFromFile<T>(string filePath)
		{
			string jsonString = File.ReadAllText(filePath, Encoding.Default);
			return JsonConvert.DeserializeObject<T>(jsonString)!;
		}
	}

	public static class PlanHelper
	{
		public static string CreateAddress(string address, int row)
		{
			return $"{address}!A{row + 1}";
		}
	}

	public interface IHelperService
	{
		AutomationConfig ReadAutomationConfig(AppConfig cfg);
		void WriteAutomationResult(AppConfig cfg, TestPlan param);
		void WriteApproval<T>(AppConfig cfg, T param, string moduleName) where T : class;
	}

	public class HelperService : IHelperService
	{
		public Dictionary<int, PVReimbursementModels> _dictPVReimbursement = [];
		public Dictionary<int, PVSettlementModels> _dictPVSettlement = [];
		public Dictionary<int, PVVendorModels> _dictPVVendor = [];
		private static void DirectoryCheck(string filePath)
		{
			var directory = Path.GetDirectoryName(filePath);
			if (!Directory.Exists(directory))
				Directory.CreateDirectory(directory!);

		}		

		public AutomationConfig ReadAutomationConfig(AppConfig cfg)
		{
			AutomationConfig result = new();
			try
			{
				Console.WriteLine($"Starting ReadAutomationConfig");

				using (FileStream fs = new(cfg.ExcelConfigPath, FileMode.Open, FileAccess.Read))
				{
					using (IWorkbook wb = new XSSFWorkbook(fs))
					{
						result.Users = ReadUsers(wb);
						result.TestPlans = ReadTestPlans(wb);
						var newTestPlans = result.TestPlans.Where(x => string.IsNullOrEmpty(x.Status)).ToList();

						#region Vendor Database
						List<VendorDatabase> vendorDatabases = ReadVendorDatabase(wb, newTestPlans);
						List<VendorContactPersonParam> vendorContactPersonParams = ReadVendorContactPersonParam(wb);
						List<VendorBankAccountParam> vendorBankAccountParams = ReadVendorBankAccountParam(wb);
						List<VendorBusinessNameParam> vendorBusinessNameParams = ReadVendorBusinessNameParam(wb);
						List<VendorDatabase> newVendorDatabases = [];

						foreach (VendorDatabase item in vendorDatabases)
						{
							item.ContactPersons = vendorContactPersonParams.Where(x => x.TestCaseId.Equals(item.TestCaseId) && x.Sequence.Equals(item.Sequence) && x.DataFor.Equals(item.DataFor)).ToList();
							item.BankAccounts = vendorBankAccountParams.Where(x => x.TestCaseId.Equals(item.TestCaseId) && x.Sequence.Equals(item.Sequence) && x.DataFor.Equals(item.DataFor)).ToList();
							item.BusinessNames = vendorBusinessNameParams.Where(x => x.TestCaseId.Equals(item.TestCaseId) && x.Sequence.Equals(item.Sequence) && x.DataFor.Equals(item.DataFor)).ToList();
							newVendorDatabases.Add(item);
						}

						result.VendorDatabases = newVendorDatabases;
						result.VendorDatabaseTasks = ReadVendorDatabaseTask(wb, newTestPlans);
						#endregion

						#region Payment Voucher
						result.PVTasks = ReadPVTask(wb, newTestPlans);
						result.PVReimbursements = ReadPVReimbursement(wb, newTestPlans);
						result.PVAdvances = ReadPVAdvance(wb, newTestPlans);
						result.PVSettlements = ReadPVSettlement(wb, newTestPlans);
						result.PVRDDs = ReadPVRDD(wb, newTestPlans);
						result.PVSalesForces = ReadPVSalesForce(wb, newTestPlans);
						result.PVVendors = ReadPVVendor(wb, newTestPlans);
						result.PVPolicyHolders = ReadPVPolicyHolder(wb, newTestPlans);
						result.PVInternalBanks = ReadPVInternalBank(wb, newTestPlans);
						result.PVSpecialCustomers = ReadPVSpecialCustomer(wb, newTestPlans);
						result.PVAgencyCommissions = ReadPVAgencyCommission(wb, newTestPlans);
						result.PVBancassuranceBMIs = PVBancassuranceBMI(wb, newTestPlans);
						result.PVBancassuranceBTNs = PVBancassuranceBTN(wb, newTestPlans);
						#endregion

						#region Bank Confirmation
						result.RejectBankConfirmations = RejectBankConfirmations(wb, newTestPlans);
						#endregion

						wb.Close();
						fs.Close();
					}
				}
				Console.WriteLine($"End ReadAutomationConfig");
				return result;
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Fail ReadAutomationConfig : \n - {ex.Message}");
				return result;
			}
		}

		#region Read
		private List<BankConfirmationModels> RejectBankConfirmations(IWorkbook wb, List<TestPlan> plans)
		{
			List<BankConfirmationModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.RejectBankConfirmation).ToList();
			try
			{
				string workSheetName = SheetConstant.RejectBankConfirmation;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexRequestNo = colIndexId + 1;
				int colIndexTransferDate = colIndexRequestNo + 1;
				int colIndexBankSource = colIndexTransferDate + 1;
				int colIndexRate = colIndexBankSource + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderRequestNo = firstRow.GetCell(colIndexRequestNo);			
				ICell cellHeaderTransferDate = firstRow.GetCell(colIndexTransferDate);
				ICell cellHeaderBankSource = firstRow.GetCell(colIndexBankSource);
				ICell cellHeaderRate = firstRow.GetCell(colIndexRate);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderRequestNo, workSheetName, colIndexRequestNo + 1, "Request No");		
				this.ValidateCellHeaderName(cellHeaderTransferDate, workSheetName, colIndexTransferDate + 1, "Transfer Date");
				this.ValidateCellHeaderName(cellHeaderBankSource, workSheetName, colIndexBankSource + 1, "Bank Source");
				this.ValidateCellHeaderName(cellHeaderRate, workSheetName, colIndexRate + 1, "Rate");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRate))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						BankConfirmationModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							RequestNo = this.TryGetString(currentRow.GetCell(colIndexRequestNo), workSheetName, cellHeaderRequestNo.StringCellValue, rowNum),
							TransferDate = this.TryGetString(currentRow.GetCell(colIndexTransferDate), workSheetName, cellHeaderTransferDate.StringCellValue, rowNum),
							BankSource = this.TryGetString(currentRow.GetCell(colIndexBankSource), workSheetName, cellHeaderBankSource.StringCellValue, rowNum),
							Rate = this.TryGetString(currentRow.GetCell(colIndexRate), workSheetName, cellHeaderRate.StringCellValue, rowNum)
						};
						
						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.RejectBankConfirmation} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVTask> ReadPVTask(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVTask> result = [];
			plans = plans.Where(x => x.TestCase == TestCaseConstant.Approval && (x.ModuleName != ModuleNameConstant.VendorDatabase || x.ModuleName != ModuleNameConstant.RejectBankConfirmation)).ToList();
			try
			{
				string workSheetName = SheetConstant.PVTask;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexSeq = colIndexId + 1;
				int colIndexActor = colIndexSeq + 1;
				int colIndexAction = colIndexActor + 1;
				int colIndexPayment = colIndexAction + 1;
				int colIndexReturnTo = colIndexPayment + 1;
				int colIndexNotes = colIndexReturnTo + 1;
				int colIndexResult = colIndexNotes + 1;
				int colIndexScreenshot = colIndexResult + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderActor = firstRow.GetCell(colIndexActor);
				ICell cellHeaderAction = firstRow.GetCell(colIndexAction);
				ICell cellHeaderPayment = firstRow.GetCell(colIndexPayment);
				ICell cellHeaderReturnTo = firstRow.GetCell(colIndexReturnTo);
				ICell cellHeaderNotes = firstRow.GetCell(colIndexNotes);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderReturnTo, workSheetName, colIndexReturnTo + 1, "Return To");
				this.ValidateCellHeaderName(cellHeaderNotes, workSheetName, colIndexNotes + 1, "Notes");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexScreenshot))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVTask param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							Sequence = int.Parse(this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum)),
							Actor = this.TryGetString(currentRow.GetCell(colIndexActor), workSheetName, cellHeaderActor.StringCellValue, rowNum),
							Action = this.TryGetString(currentRow.GetCell(colIndexAction), workSheetName, cellHeaderAction.StringCellValue, rowNum),
							IsPayment = this.TryGetBoolean(currentRow.GetCell(colIndexPayment), workSheetName, cellHeaderPayment.StringCellValue, rowNum),
							ReturnTo = this.TryGetString(currentRow.GetCell(colIndexReturnTo), workSheetName, cellHeaderReturnTo.StringCellValue, rowNum),
							Notes = this.TryGetString(currentRow.GetCell(colIndexNotes), workSheetName, cellHeaderNotes.StringCellValue, rowNum)
						};
						
						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVTask} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVSpecialCustomerModels> ReadPVSpecialCustomer(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVSpecialCustomerModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVSpecialCustomer).ToList();
			try
			{
				string workSheetName = SheetConstant.PVSpecialCustomer;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexRequestorName = colIndexPaymentPurpose + 1;
				int colIndexAttachmentPath = colIndexRequestorName + 1;
				int colIndexAttachmentDesc = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDesc + 1;
				int colIndexPayee = colIndexRemarks + 1;
				int colIndexPayeeType = colIndexPayee + 1;
				int colIndexCountry = colIndexPayeeType + 1;
				int colIndexIsTaxable = colIndexCountry + 1;
				int colIndexSKPPPath = colIndexIsTaxable + 1;
				int colIndexIsCoD = colIndexSKPPPath + 1;
				int colIndexCoDPath = colIndexIsCoD + 1;
				int colIndexCoDExpiryDate = colIndexCoDPath + 1;
				int colIndexIsUWTax = colIndexCoDExpiryDate + 1;
				int colIndexTaxIdNo = colIndexIsUWTax + 1;
				int colIndexTaxIdName = colIndexTaxIdNo + 1;
				int colIndexTaxIdAddress = colIndexTaxIdName + 1;
				int colIndexTaxIdAttachmentPath = colIndexTaxIdAddress + 1;
				int colIndexBankName = colIndexTaxIdAttachmentPath + 1;
				int colIndexBankAccountNo = colIndexBankName + 1;
				int colIndexBankAccountName = colIndexBankAccountNo + 1;
				int colIndexBranch = colIndexBankAccountName + 1;
				int colIndexBussinessName = colIndexBranch + 1;
				int colIndexCostCenter = colIndexBussinessName + 1;
				int colIndexLongDesc = colIndexCostCenter + 1;
				int colIndexShortDesc = colIndexLongDesc + 1;
				int colIndexGrossAmount = colIndexShortDesc + 1;
				int colIndexCurrency = colIndexGrossAmount + 1;
				int colIndexIsTaxExemption = colIndexCurrency + 1;
				int colIndexSKBPath = colIndexIsTaxExemption + 1;
				int colIndexDetailAttachmentPath = colIndexSKBPath + 1;
				int colIndexDetailRemarks = colIndexDetailAttachmentPath + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderRequestorName = firstRow.GetCell(colIndexRequestorName);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDesc = firstRow.GetCell(colIndexAttachmentDesc);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderPayee = firstRow.GetCell(colIndexPayee);
				ICell cellHeaderPayeeType = firstRow.GetCell(colIndexPayeeType);
				ICell cellHeaderCountry = firstRow.GetCell(colIndexCountry);
				ICell cellHeaderIsTaxable = firstRow.GetCell(colIndexIsTaxable);
				ICell cellHeaderSKPPPath = firstRow.GetCell(colIndexSKPPPath);
				ICell cellHeaderIsCoD = firstRow.GetCell(colIndexIsCoD);
				ICell cellHeaderCoDPath = firstRow.GetCell(colIndexCoDPath);
				ICell cellHeaderCoDExpiryDate = firstRow.GetCell(colIndexCoDExpiryDate);
				ICell cellHeaderIsUWTax = firstRow.GetCell(colIndexIsUWTax);
				ICell cellHeaderTaxIdNo = firstRow.GetCell(colIndexTaxIdNo);
				ICell cellHeaderTaxIdName = firstRow.GetCell(colIndexTaxIdName);
				ICell cellHeaderTaxIdAddress = firstRow.GetCell(colIndexTaxIdAddress);
				ICell cellHeaderTaxIdAttachmentPath = firstRow.GetCell(colIndexTaxIdAttachmentPath);
				ICell cellHeaderBankName = firstRow.GetCell(colIndexBankName);
				ICell cellHeaderBankAccountNo = firstRow.GetCell(colIndexBankAccountNo);
				ICell cellHeaderBankAccountName = firstRow.GetCell(colIndexBankAccountName);
				ICell cellHeaderBranch = firstRow.GetCell(colIndexBranch);
				ICell cellHeaderBussinessName = firstRow.GetCell(colIndexBussinessName);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderLongDesc = firstRow.GetCell(colIndexLongDesc);
				ICell cellHeaderShortDesc = firstRow.GetCell(colIndexShortDesc);
				ICell cellHeaderGrossAmount = firstRow.GetCell(colIndexGrossAmount);
				ICell cellHeaderCurrency = firstRow.GetCell(colIndexCurrency);
				ICell cellHeaderIsTaxExemption = firstRow.GetCell(colIndexIsTaxExemption);
				ICell cellHeadeSKBPath = firstRow.GetCell(colIndexSKBPath);
				ICell cellHeadeDetailAttachmentPath = firstRow.GetCell(colIndexDetailAttachmentPath);
				ICell cellHeadeDetailRemarks = firstRow.GetCell(colIndexDetailRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderRequestorName, workSheetName, colIndexRequestorName + 1, "Requestor Name");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDesc, workSheetName, colIndexAttachmentDesc + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderPayee, workSheetName, colIndexPayee + 1, "Payee");
				this.ValidateCellHeaderName(cellHeaderCountry, workSheetName, colIndexCountry + 1, "Country");
				this.ValidateCellHeaderName(cellHeaderIsTaxable, workSheetName, colIndexIsTaxable + 1, "Is Taxable");
				this.ValidateCellHeaderName(cellHeaderSKPPPath, workSheetName, colIndexSKPPPath + 1, "SPPK Path");
				this.ValidateCellHeaderName(cellHeaderIsCoD, workSheetName, colIndexIsCoD + 1, "Is CoD");
				this.ValidateCellHeaderName(cellHeaderCoDPath, workSheetName, colIndexCoDPath + 1, "CoD Path");
				this.ValidateCellHeaderName(cellHeaderCoDExpiryDate, workSheetName, colIndexCoDExpiryDate + 1, "CoD Expiry Date");
				this.ValidateCellHeaderName(cellHeaderIsUWTax, workSheetName, colIndexIsUWTax + 1, "Is UW Tax");
				this.ValidateCellHeaderName(cellHeaderTaxIdNo, workSheetName, colIndexTaxIdNo + 1, "Tax Id No");
				this.ValidateCellHeaderName(cellHeaderTaxIdName, workSheetName, colIndexTaxIdName + 1, "Tax Id Name");
				this.ValidateCellHeaderName(cellHeaderTaxIdAddress, workSheetName, colIndexTaxIdAddress + 1, "Tax Id Address");
				this.ValidateCellHeaderName(cellHeaderTaxIdAttachmentPath, workSheetName, colIndexTaxIdAttachmentPath + 1, "Tax Id Attachment Path");
				this.ValidateCellHeaderName(cellHeaderBankName, workSheetName, colIndexBankName + 1, "Bank Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountNo, workSheetName, colIndexBankAccountNo + 1, "Bank Account No");
				this.ValidateCellHeaderName(cellHeaderBankAccountName, workSheetName, colIndexBankAccountName + 1, "Bank Account Name");
				this.ValidateCellHeaderName(cellHeaderBranch, workSheetName, colIndexBranch + 1, "Branch");
				this.ValidateCellHeaderName(cellHeaderBussinessName, workSheetName, colIndexBussinessName + 1, "Bussiness Name");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center");
				this.ValidateCellHeaderName(cellHeaderLongDesc, workSheetName, colIndexLongDesc + 1, "Long Description");
				this.ValidateCellHeaderName(cellHeaderShortDesc, workSheetName, colIndexShortDesc + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderGrossAmount, workSheetName, colIndexGrossAmount + 1, "Gross Amount");
				this.ValidateCellHeaderName(cellHeaderCurrency, workSheetName, colIndexCurrency + 1, "Currency");
				this.ValidateCellHeaderName(cellHeaderIsTaxExemption, workSheetName, colIndexIsTaxExemption + 1, "Is Tax Exemption");
				this.ValidateCellHeaderName(cellHeadeSKBPath, workSheetName, colIndexSKBPath + 1, "SKB Path");
				this.ValidateCellHeaderName(cellHeadeDetailAttachmentPath, workSheetName, colIndexDetailAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeadeDetailRemarks, workSheetName, colIndexDetailRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexDetailRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVSpecialCustomerModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							RequestorName = this.TryGetString(currentRow.GetCell(colIndexRequestorName), workSheetName, cellHeaderRequestorName.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDesc), workSheetName, cellHeaderAttachmentDesc.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum),
							Payee = this.TryGetString(currentRow.GetCell(colIndexPayee), workSheetName, cellHeaderPayee.StringCellValue, rowNum),
							PayeeType = this.TryGetString(currentRow.GetCell(colIndexPayeeType), workSheetName, cellHeaderPayeeType.StringCellValue, rowNum),
							Country = this.TryGetString(currentRow.GetCell(colIndexCountry), workSheetName, cellHeaderCountry.StringCellValue, rowNum),
							IsTaxable = this.TryGetBoolean(currentRow.GetCell(colIndexIsTaxable), workSheetName, cellHeaderIsTaxable.StringCellValue, rowNum),
							SKPPPath = this.TryGetString(currentRow.GetCell(colIndexSKPPPath), workSheetName, cellHeaderSKPPPath.StringCellValue, rowNum),
							IsCoD = this.TryGetBoolean(currentRow.GetCell(colIndexIsCoD), workSheetName, cellHeaderIsCoD.StringCellValue, rowNum),
							CoDPath = this.TryGetString(currentRow.GetCell(colIndexCoDPath), workSheetName, cellHeaderCoDPath.StringCellValue, rowNum),
							CoDExpiryDate = this.TryGetString(currentRow.GetCell(colIndexCoDExpiryDate), workSheetName, cellHeaderCoDExpiryDate.StringCellValue, rowNum),
							IsUWTax = this.TryGetBoolean(currentRow.GetCell(colIndexIsUWTax), workSheetName, cellHeaderIsUWTax.StringCellValue, rowNum),
							TaxIdNo = this.TryGetString(currentRow.GetCell(colIndexTaxIdNo), workSheetName, cellHeaderTaxIdNo.StringCellValue, rowNum),
							TaxIdName = this.TryGetString(currentRow.GetCell(colIndexTaxIdName), workSheetName, cellHeaderTaxIdName.StringCellValue, rowNum),
							TaxIdAddress = this.TryGetString(currentRow.GetCell(colIndexTaxIdAddress), workSheetName, cellHeaderTaxIdAddress.StringCellValue, rowNum),
							TaxIdAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexTaxIdAttachmentPath), workSheetName, cellHeaderTaxIdAttachmentPath.StringCellValue, rowNum),
							BankName = this.TryGetString(currentRow.GetCell(colIndexBankName), workSheetName, cellHeaderBankName.StringCellValue, rowNum),
							BankAccountNo = this.TryGetString(currentRow.GetCell(colIndexBankAccountNo), workSheetName, cellHeaderBankAccountNo.StringCellValue, rowNum),
							BankAccountName = this.TryGetString(currentRow.GetCell(colIndexBankAccountName), workSheetName, cellHeaderBankAccountName.StringCellValue, rowNum),
							Branch = this.TryGetString(currentRow.GetCell(colIndexBranch), workSheetName, cellHeaderBranch.StringCellValue, rowNum),

							Detail = new()
							{
								BussinessName = this.TryGetString(currentRow.GetCell(colIndexBussinessName), workSheetName, cellHeaderBussinessName.StringCellValue, rowNum),
								CostCenter = this.TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
								LongDesc = this.TryGetString(currentRow.GetCell(colIndexLongDesc), workSheetName, cellHeaderLongDesc.StringCellValue, rowNum),
								ShortDesc = this.TryGetString(currentRow.GetCell(colIndexShortDesc), workSheetName, cellHeaderShortDesc.StringCellValue, rowNum),
								GrossAmount = this.TryGetString(currentRow.GetCell(colIndexGrossAmount), workSheetName, cellHeaderGrossAmount.StringCellValue, rowNum),
								Currency = this.TryGetString(currentRow.GetCell(colIndexCurrency), workSheetName, cellHeaderCurrency.StringCellValue, rowNum),
								IsTaxExemption = this.TryGetBoolean(currentRow.GetCell(colIndexIsTaxExemption), workSheetName, cellHeaderIsTaxExemption.StringCellValue, rowNum),
								SKBPath = this.TryGetString(currentRow.GetCell(colIndexSKBPath), workSheetName, cellHeadeSKBPath.StringCellValue, rowNum),
								AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDetailAttachmentPath), workSheetName, cellHeadeDetailRemarks.StringCellValue, rowNum),
								Remarks = this.TryGetString(currentRow.GetCell(colIndexDetailRemarks), workSheetName, cellHeadeDetailRemarks.StringCellValue, rowNum)
							}
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVSpecialCustomer} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVInternalBankModels> ReadPVInternalBank(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVInternalBankModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVInternalBank).ToList();
			try
			{
				string workSheetName = SheetConstant.PVInternalBank;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexAttachmentPath = colIndexPaymentPurpose + 1;
				int colIndexAttachmentDesc = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDesc + 1;
				int colIndexFromBankAccountNo = colIndexRemarks + 1;
				int colIndexToBankAccountNo = colIndexFromBankAccountNo + 1;
				int colIndexCostCenter = colIndexToBankAccountNo + 1;
				int colIndexLongDesc = colIndexCostCenter + 1;
				int colIndexShortDesc = colIndexLongDesc + 1;
				int colIndexCurrency = colIndexShortDesc + 1;
				int colIndexAmount = colIndexCurrency + 1;
				int colIndexDetailAttachmentPath = colIndexAmount + 1;
				int colIndexDetailRemarks = colIndexDetailAttachmentPath + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDesc = firstRow.GetCell(colIndexAttachmentDesc);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderFromBankAccountNo = firstRow.GetCell(colIndexFromBankAccountNo);
				ICell cellHeaderToBankAccountNo = firstRow.GetCell(colIndexToBankAccountNo);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderLongDesc = firstRow.GetCell(colIndexLongDesc);
				ICell cellHeaderShortDesc = firstRow.GetCell(colIndexShortDesc);
				ICell cellHeaderCurrency = firstRow.GetCell(colIndexCurrency);
				ICell cellHeaderAmount = firstRow.GetCell(colIndexAmount);
				ICell cellHeaderDetailAttachmentPath = firstRow.GetCell(colIndexDetailAttachmentPath);
				ICell cellHeadeDetailRemarks = firstRow.GetCell(colIndexDetailRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDesc, workSheetName, colIndexAttachmentDesc + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderFromBankAccountNo, workSheetName, colIndexFromBankAccountNo + 1, "From Bank Account No");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center");
				this.ValidateCellHeaderName(cellHeaderLongDesc, workSheetName, colIndexLongDesc + 1, "Long Description");
				this.ValidateCellHeaderName(cellHeaderShortDesc, workSheetName, colIndexShortDesc + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderCurrency, workSheetName, colIndexCurrency + 1, "Currency");
				this.ValidateCellHeaderName(cellHeaderAmount, workSheetName, colIndexAmount + 1, "Amount");
				this.ValidateCellHeaderName(cellHeaderDetailAttachmentPath, workSheetName, colIndexDetailAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeadeDetailRemarks, workSheetName, colIndexDetailRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexDetailRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVInternalBankModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDesc), workSheetName, cellHeaderAttachmentDesc.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum),
							FromBankAccountNo = this.TryGetString(currentRow.GetCell(colIndexFromBankAccountNo), workSheetName, cellHeaderFromBankAccountNo.StringCellValue, rowNum),
							ToBankAccountNo = this.TryGetString(currentRow.GetCell(colIndexToBankAccountNo), workSheetName, cellHeaderToBankAccountNo.StringCellValue, rowNum),
							CostCenter = this.TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
							LongDesc = this.TryGetString(currentRow.GetCell(colIndexLongDesc), workSheetName, cellHeaderLongDesc.StringCellValue, rowNum),
							ShortDesc = this.TryGetString(currentRow.GetCell(colIndexShortDesc), workSheetName, cellHeaderShortDesc.StringCellValue, rowNum),
							Currency = this.TryGetString(currentRow.GetCell(colIndexCurrency), workSheetName, cellHeaderCurrency.StringCellValue, rowNum),
							Amount = this.TryGetString(currentRow.GetCell(colIndexAmount), workSheetName, cellHeaderAmount.StringCellValue, rowNum),
							DetailAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDetailAttachmentPath), workSheetName, cellHeaderDetailAttachmentPath.StringCellValue, rowNum),
							DetailRemarks = this.TryGetString(currentRow.GetCell(colIndexDetailRemarks), workSheetName, cellHeadeDetailRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVInternalBank} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVPolicyHolderModels> ReadPVPolicyHolder(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVPolicyHolderModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVPolicyHolder).ToList();
			try
			{
				string workSheetName = SheetConstant.PVPolicyHolder;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexRequestorName = colIndexPaymentPurpose + 1;
				int colIndexAttachmentPath = colIndexRequestorName + 1;
				int colIndexAttachmentDesc = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDesc + 1;
				int colIndexPayee = colIndexRemarks + 1;
				int colIndexPolicyNo = colIndexPayee + 1;
				int colIndexSPAJNo = colIndexPolicyNo + 1;
				int colIndexBussinessName = colIndexSPAJNo + 1;
				int colIndexLongDesc = colIndexBussinessName + 1;
				int colIndexShortDesc = colIndexLongDesc + 1;
				int colIndexProduct = colIndexShortDesc + 1;
				int colIndexDistribution = colIndexProduct + 1;
				int colIndexCountry = colIndexDistribution + 1;
				int colIndexBankName = colIndexCountry + 1;
				int colIndexBankAccountNo = colIndexBankName + 1;
				int colIndexBankAccountName = colIndexBankAccountNo + 1;
				int colIndexBranch = colIndexBankAccountName + 1;
				int colIndexCurrency = colIndexBranch + 1;
				int colIndexGrossAmount = colIndexCurrency + 1;
				int colIndexTaxAmount = colIndexGrossAmount + 1;
				int colIndexChargesAdmin = colIndexTaxAmount + 1;
				int colIndexChargesMedical = colIndexChargesAdmin + 1;
				int colIndexForPremiumPayment = colIndexChargesMedical + 1;
				int colIndexDetailAttachmentPath = colIndexForPremiumPayment + 1;
				int colIndexTaxIdNo = colIndexDetailAttachmentPath + 1;
				int colIndexTaxIdName = colIndexTaxIdNo + 1;
				int colIndexTaxIdAddress = colIndexTaxIdName + 1;
				int colIndexTaxIdAttachmentPath = colIndexTaxIdAddress + 1;
				int colIndexDetailRemarks = colIndexTaxIdAttachmentPath + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderRequestorName = firstRow.GetCell(colIndexRequestorName);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDesc = firstRow.GetCell(colIndexAttachmentDesc);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderPayee = firstRow.GetCell(colIndexPayee);
				ICell cellHeaderPolicyNo = firstRow.GetCell(colIndexPolicyNo);
				ICell cellHeaderSPAJNo = firstRow.GetCell(colIndexSPAJNo);
				ICell cellHeaderBussinessName = firstRow.GetCell(colIndexBussinessName);
				ICell cellHeaderLongDesc = firstRow.GetCell(colIndexLongDesc);
				ICell cellHeaderShortDesc = firstRow.GetCell(colIndexShortDesc);
				ICell cellHeaderProduct = firstRow.GetCell(colIndexProduct);
				ICell cellHeaderDistribution = firstRow.GetCell(colIndexDistribution);
				ICell cellHeaderCountry = firstRow.GetCell(colIndexCountry);
				ICell cellHeaderBankName = firstRow.GetCell(colIndexBankName);
				ICell cellHeaderBankAccountNo = firstRow.GetCell(colIndexBankAccountNo);
				ICell cellHeaderBankAccountName = firstRow.GetCell(colIndexBankAccountName);
				ICell cellHeaderBranch = firstRow.GetCell(colIndexBranch);
				ICell cellHeaderCurrency = firstRow.GetCell(colIndexCurrency);
				ICell cellHeaderGrossAmount = firstRow.GetCell(colIndexGrossAmount);
				ICell cellHeaderTaxAmount = firstRow.GetCell(colIndexTaxAmount);
				ICell cellHeaderChargesAdmin = firstRow.GetCell(colIndexChargesAdmin);
				ICell cellHeaderChargesMedical = firstRow.GetCell(colIndexChargesMedical);
				ICell cellHeaderForPremiumPayment = firstRow.GetCell(colIndexForPremiumPayment);
				ICell cellHeaderDetailAttachmentPath = firstRow.GetCell(colIndexDetailAttachmentPath);
				ICell cellHeaderTaxIdNo = firstRow.GetCell(colIndexTaxIdNo);
				ICell cellHeaderTaxIdName = firstRow.GetCell(colIndexTaxIdName);
				ICell cellHeaderTaxIdAddress = firstRow.GetCell(colIndexTaxIdAddress);
				ICell cellHeaderTaxIdAttachmentPath = firstRow.GetCell(colIndexTaxIdAttachmentPath);
				ICell cellHeadeDetailRemarks = firstRow.GetCell(colIndexDetailRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderRequestorName, workSheetName, colIndexRequestorName + 1, "Requestor Name");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDesc, workSheetName, colIndexAttachmentDesc + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderPayee, workSheetName, colIndexPayee + 1, "Payee");
				this.ValidateCellHeaderName(cellHeaderSPAJNo, workSheetName, colIndexSPAJNo + 1, "SPAJ No");
				this.ValidateCellHeaderName(cellHeaderBussinessName, workSheetName, colIndexBussinessName + 1, "Bussiness Name");
				this.ValidateCellHeaderName(cellHeaderLongDesc, workSheetName, colIndexLongDesc + 1, "Long Description");
				this.ValidateCellHeaderName(cellHeaderShortDesc, workSheetName, colIndexShortDesc + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderProduct, workSheetName, colIndexProduct + 1, "Product");
				this.ValidateCellHeaderName(cellHeaderDistribution, workSheetName, colIndexDistribution + 1, "Distribution");
				this.ValidateCellHeaderName(cellHeaderCountry, workSheetName, colIndexCountry + 1, "Country");
				this.ValidateCellHeaderName(cellHeaderBankName, workSheetName, colIndexBankName + 1, "Bank Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountNo, workSheetName, colIndexBankAccountNo + 1, "Bank Account No");
				this.ValidateCellHeaderName(cellHeaderBankAccountName, workSheetName, colIndexBankAccountName + 1, "Bank Account Name");
				this.ValidateCellHeaderName(cellHeaderBranch, workSheetName, colIndexBranch + 1, "Branch");
				this.ValidateCellHeaderName(cellHeaderCurrency, workSheetName, colIndexCurrency + 1, "Currency");
				this.ValidateCellHeaderName(cellHeaderGrossAmount, workSheetName, colIndexGrossAmount + 1, "Gross Amount");
				this.ValidateCellHeaderName(cellHeaderTaxAmount, workSheetName, colIndexTaxAmount + 1, "Tax Amount");
				this.ValidateCellHeaderName(cellHeaderChargesAdmin, workSheetName, colIndexChargesAdmin + 1, "Charges Admin");
				this.ValidateCellHeaderName(cellHeaderChargesMedical, workSheetName, colIndexChargesMedical + 1, "Charges Medical");
				this.ValidateCellHeaderName(cellHeaderForPremiumPayment, workSheetName, colIndexForPremiumPayment + 1, "For Premium Payment");
				this.ValidateCellHeaderName(cellHeaderDetailAttachmentPath, workSheetName, colIndexDetailAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderTaxIdNo, workSheetName, colIndexTaxIdNo + 1, "Tax Id No");
				this.ValidateCellHeaderName(cellHeaderTaxIdName, workSheetName, colIndexTaxIdName + 1, "Tax Id Name");
				this.ValidateCellHeaderName(cellHeaderTaxIdAddress, workSheetName, colIndexTaxIdAddress + 1, "Tax Id Address");
				this.ValidateCellHeaderName(cellHeaderTaxIdAttachmentPath, workSheetName, colIndexTaxIdAttachmentPath + 1, "Tax Id Attachment Path");
				this.ValidateCellHeaderName(cellHeadeDetailRemarks, workSheetName, colIndexDetailRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexDetailRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVPolicyHolderModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							RequestorName = this.TryGetString(currentRow.GetCell(colIndexRequestorName), workSheetName, cellHeaderRequestorName.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDesc), workSheetName, cellHeaderAttachmentDesc.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum),
							Payee = this.TryGetString(currentRow.GetCell(colIndexPayee), workSheetName, cellHeaderPayee.StringCellValue, rowNum),
							PolicyNo = this.TryGetString(currentRow.GetCell(colIndexPolicyNo), workSheetName, cellHeaderPolicyNo.StringCellValue, rowNum),
							SPAJNo = this.TryGetString(currentRow.GetCell(colIndexSPAJNo), workSheetName, cellHeaderSPAJNo.StringCellValue, rowNum),
							BussinessName = this.TryGetString(currentRow.GetCell(colIndexBussinessName), workSheetName, cellHeaderBussinessName.StringCellValue, rowNum),
							LongDesc = this.TryGetString(currentRow.GetCell(colIndexLongDesc), workSheetName, cellHeaderLongDesc.StringCellValue, rowNum),
							ShortDesc = this.TryGetString(currentRow.GetCell(colIndexShortDesc), workSheetName, cellHeaderShortDesc.StringCellValue, rowNum),
							Product = this.TryGetString(currentRow.GetCell(colIndexProduct), workSheetName, cellHeaderProduct.StringCellValue, rowNum),
							Distribution = this.TryGetString(currentRow.GetCell(colIndexDistribution), workSheetName, cellHeaderDistribution.StringCellValue, rowNum),
							Country = this.TryGetString(currentRow.GetCell(colIndexCountry), workSheetName, cellHeaderCountry.StringCellValue, rowNum),
							BankName = this.TryGetString(currentRow.GetCell(colIndexBankName), workSheetName, cellHeaderBankName.StringCellValue, rowNum),
							BankAccountNo = this.TryGetString(currentRow.GetCell(colIndexBankAccountNo), workSheetName, cellHeaderBankAccountNo.StringCellValue, rowNum),
							BankAccountName = this.TryGetString(currentRow.GetCell(colIndexBankAccountName), workSheetName, cellHeaderBankAccountName.StringCellValue, rowNum),
							Branch = this.TryGetString(currentRow.GetCell(colIndexBranch), workSheetName, cellHeaderBranch.StringCellValue, rowNum),
							Currency = this.TryGetString(currentRow.GetCell(colIndexCurrency), workSheetName, cellHeaderCurrency.StringCellValue, rowNum),
							GrossAmount = this.TryGetString(currentRow.GetCell(colIndexGrossAmount), workSheetName, cellHeaderGrossAmount.StringCellValue, rowNum),
							TaxAmount = this.TryGetString(currentRow.GetCell(colIndexTaxAmount), workSheetName, cellHeaderTaxAmount.StringCellValue, rowNum),
							ChargesAdmin = this.TryGetString(currentRow.GetCell(colIndexChargesAdmin), workSheetName, cellHeaderChargesAdmin.StringCellValue, rowNum),
							ChargesMedical = this.TryGetString(currentRow.GetCell(colIndexChargesMedical), workSheetName, cellHeaderChargesMedical.StringCellValue, rowNum),
							ForPremiumPayment = this.TryGetString(currentRow.GetCell(colIndexForPremiumPayment), workSheetName, cellHeaderForPremiumPayment.StringCellValue, rowNum),
							DetailAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDetailAttachmentPath), workSheetName, cellHeaderDetailAttachmentPath.StringCellValue, rowNum),
							TaxIdNo = this.TryGetString(currentRow.GetCell(colIndexTaxIdNo), workSheetName, cellHeaderTaxIdNo.StringCellValue, rowNum),
							TaxIdName = this.TryGetString(currentRow.GetCell(colIndexTaxIdName), workSheetName, cellHeaderTaxIdName.StringCellValue, rowNum),
							TaxIdAddress = this.TryGetString(currentRow.GetCell(colIndexTaxIdAddress), workSheetName, cellHeaderTaxIdAddress.StringCellValue, rowNum),
							TaxIdAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexTaxIdAttachmentPath), workSheetName, cellHeaderTaxIdAttachmentPath.StringCellValue, rowNum),
							DetailRemarks = this.TryGetString(currentRow.GetCell(colIndexDetailRemarks), workSheetName, cellHeadeDetailRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVPolicyHolder} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVBancassuranceBTNModels> PVBancassuranceBTN(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVBancassuranceBTNModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVBancassuranceBTN).ToList();
			try
			{
				string workSheetName = SheetConstant.PVBancassuranceBTN;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexTaxReportWABTNPath = colIndexPaymentPurpose + 1;
				int colIndexTemplateAdjustmentWABTNPath = colIndexTaxReportWABTNPath + 1;
				int colIndexTemplateFeebaseBTNPath = colIndexTemplateAdjustmentWABTNPath + 1;
				int colIndexDetailBTNPath = colIndexTemplateFeebaseBTNPath + 1;
				int colIndexCostCenterName = colIndexDetailBTNPath + 1;
				int colIndexAttachmentPath = colIndexCostCenterName + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderTaxReportWABTNPath = firstRow.GetCell(colIndexTaxReportWABTNPath);
				ICell cellHeaderTemplateAdjustmentWABTNPath = firstRow.GetCell(colIndexTemplateAdjustmentWABTNPath);
				ICell cellHeaderTemplateFeebaseBTNPath = firstRow.GetCell(colIndexTemplateFeebaseBTNPath);
				ICell cellHeaderDetailBTNPath = firstRow.GetCell(colIndexDetailBTNPath);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenterName);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderTaxReportWABTNPath, workSheetName, colIndexTaxReportWABTNPath + 1, "Tax Report WMA BTN Path");
				this.ValidateCellHeaderName(cellHeaderTemplateAdjustmentWABTNPath, workSheetName, colIndexTemplateAdjustmentWABTNPath + 1, "Template Adjustment WMA BTN Path");
				this.ValidateCellHeaderName(cellHeaderTemplateFeebaseBTNPath, workSheetName, colIndexTemplateFeebaseBTNPath + 1, "Template Feebase BTN Path");
				this.ValidateCellHeaderName(cellHeaderDetailBTNPath, workSheetName, colIndexDetailBTNPath + 1, "Detail BTN Path");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenterName + 1, "Cost Center Name");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVBancassuranceBTNModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							TaxReportWMABTNPath = this.TryGetString(currentRow.GetCell(colIndexTaxReportWABTNPath), workSheetName, cellHeaderTaxReportWABTNPath.StringCellValue, rowNum),
							TemplateAdjustmentWMABTNPath = this.TryGetString(currentRow.GetCell(colIndexTemplateAdjustmentWABTNPath), workSheetName, cellHeaderTemplateAdjustmentWABTNPath.StringCellValue, rowNum),
							TemplateFeebaseBTNPath = this.TryGetString(currentRow.GetCell(colIndexTemplateFeebaseBTNPath), workSheetName, cellHeaderTemplateFeebaseBTNPath.StringCellValue, rowNum),
							DetailBTNPath = this.TryGetString(currentRow.GetCell(colIndexDetailBTNPath), workSheetName, cellHeaderDetailBTNPath.StringCellValue, rowNum),
							CostCenterName = this.TryGetString(currentRow.GetCell(colIndexCostCenterName), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVBancassuranceBTN} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVBancassuranceBMIModels> PVBancassuranceBMI(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVBancassuranceBMIModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVBancassuranceBMI).ToList();
			try
			{
				string workSheetName = SheetConstant.PVBancassuranceBMI;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexTaxReportWABMIPath = colIndexPaymentPurpose + 1;
				int colIndexTemplateAdjustmentWABMIPath = colIndexTaxReportWABMIPath + 1;
				int colIndexTemplateFeebaseBMIPath = colIndexTemplateAdjustmentWABMIPath + 1;
				int colIndexTemplateCreditLifeBMIPath = colIndexTemplateFeebaseBMIPath + 1;
				int colIndexCreditLifeBMIPath = colIndexTemplateCreditLifeBMIPath + 1;
				int colIndexTaxCalculationBSPath = colIndexCreditLifeBMIPath + 1;
				int colIndexTemplateAdjustmentBankStaffBMIPath = colIndexTaxCalculationBSPath + 1;
				int colIndexDetailBMIPath = colIndexTemplateAdjustmentBankStaffBMIPath + 1;
				int colIndexCostCenterName = colIndexDetailBMIPath + 1;
				int colIndexAttachmentPath = colIndexCostCenterName + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderTaxReportWABMIPath = firstRow.GetCell(colIndexTaxReportWABMIPath);
				ICell cellHeaderTemplateAdjustmentWABMIPath = firstRow.GetCell(colIndexTemplateAdjustmentWABMIPath);
				ICell cellHeaderTemplateFeebaseBMIPath = firstRow.GetCell(colIndexTemplateFeebaseBMIPath);
				ICell cellHeaderTemplateCreditLifeBMIPath = firstRow.GetCell(colIndexTemplateCreditLifeBMIPath);
				ICell cellHeaderCreditLifeBMIPath = firstRow.GetCell(colIndexCreditLifeBMIPath);
				ICell cellHeaderTaxCalculationBSPath = firstRow.GetCell(colIndexTaxCalculationBSPath);
				ICell cellHeaderTemplateAdjustmentBankStaffBMIPath = firstRow.GetCell(colIndexTemplateAdjustmentBankStaffBMIPath);
				ICell cellHeaderDetailBMIPath = firstRow.GetCell(colIndexDetailBMIPath);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenterName);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderTaxReportWABMIPath, workSheetName, colIndexTaxReportWABMIPath + 1, "Tax Report WMA BMI Path");
				this.ValidateCellHeaderName(cellHeaderTemplateAdjustmentWABMIPath, workSheetName, colIndexTemplateAdjustmentWABMIPath + 1, "Template Adjustment WMA BMI Path");
				this.ValidateCellHeaderName(cellHeaderTemplateFeebaseBMIPath, workSheetName, colIndexTemplateFeebaseBMIPath + 1, "Template Feebase BMI Path");
				this.ValidateCellHeaderName(cellHeaderTemplateCreditLifeBMIPath, workSheetName, colIndexTemplateCreditLifeBMIPath + 1, "Template Credit Life BMI Path");
				this.ValidateCellHeaderName(cellHeaderCreditLifeBMIPath, workSheetName, colIndexCreditLifeBMIPath + 1, "Credit Life BMI Path");
				this.ValidateCellHeaderName(cellHeaderTaxCalculationBSPath, workSheetName, colIndexTaxCalculationBSPath + 1, "Tax Calculation BS Path");
				this.ValidateCellHeaderName(cellHeaderTemplateAdjustmentBankStaffBMIPath, workSheetName, colIndexTemplateAdjustmentBankStaffBMIPath + 1, "Template Adjustment Bank Staff BMI Path");
				this.ValidateCellHeaderName(cellHeaderDetailBMIPath, workSheetName, colIndexDetailBMIPath + 1, "Detail BMI Path");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenterName + 1, "Cost Center Name");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVBancassuranceBMIModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							TaxReportWABMIPath = this.TryGetString(currentRow.GetCell(colIndexTaxReportWABMIPath), workSheetName, cellHeaderTaxReportWABMIPath.StringCellValue, rowNum),
							TemplateAdjustmentWABMIPath = this.TryGetString(currentRow.GetCell(colIndexTemplateAdjustmentWABMIPath), workSheetName, cellHeaderTemplateAdjustmentWABMIPath.StringCellValue, rowNum),
							TemplateFeebaseBMIPath = this.TryGetString(currentRow.GetCell(colIndexTemplateFeebaseBMIPath), workSheetName, cellHeaderTemplateFeebaseBMIPath.StringCellValue, rowNum),
							TemplateCreditLifeBMIPath = this.TryGetString(currentRow.GetCell(colIndexTemplateCreditLifeBMIPath), workSheetName, cellHeaderTemplateCreditLifeBMIPath.StringCellValue, rowNum),
							CreditLifeBMIPath = this.TryGetString(currentRow.GetCell(colIndexCreditLifeBMIPath), workSheetName, cellHeaderCreditLifeBMIPath.StringCellValue, rowNum),
							TaxCalculationBSPath = this.TryGetString(currentRow.GetCell(colIndexTaxCalculationBSPath), workSheetName, cellHeaderTaxCalculationBSPath.StringCellValue, rowNum),
							TemplateAdjustmentBankStaffBMIPath = this.TryGetString(currentRow.GetCell(colIndexTemplateAdjustmentBankStaffBMIPath), workSheetName, cellHeaderTemplateAdjustmentBankStaffBMIPath.StringCellValue, rowNum),
							DetailBMIPath = this.TryGetString(currentRow.GetCell(colIndexDetailBMIPath), workSheetName, cellHeaderDetailBMIPath.StringCellValue, rowNum),
							CostCenterName = this.TryGetString(currentRow.GetCell(colIndexCostCenterName), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVBancassuranceBMI} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVAgencyCommissionModels> ReadPVAgencyCommission(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVAgencyCommissionModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVAgencyCommission).ToList();
			try
			{
				string workSheetName = SheetConstant.PVAgencyCommission;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexGAAllowancePT = colIndexPaymentPurpose + 1;
				int colIndexFinancing = colIndexGAAllowancePT + 1;
				int colIndexMGI = colIndexFinancing + 1;
				int colIndexTaxCalculation = colIndexMGI + 1;
				int colIndexFirstYearCR = colIndexTaxCalculation + 1;
				int colIndexRenewalYearCR = colIndexFirstYearCR + 1;
				int colIndexReportRecOverride = colIndexRenewalYearCR + 1;
				int colIndexADGenDetail = colIndexReportRecOverride + 1;
				int colIndexProducerBonusReport = colIndexADGenDetail + 1;
				int colIndexAALBonusReport = colIndexProducerBonusReport + 1;
				int colIndexSummaryOverride = colIndexAALBonusReport + 1;
				int colIndexGAAllowancePersonal = colIndexSummaryOverride + 1;
				int colIndexSummaryAgentPaid = colIndexGAAllowancePersonal + 1;
				int colIndexCFGManualPayment = colIndexSummaryAgentPaid + 1;
				int colIndexCostCenter = colIndexCFGManualPayment + 1;
				int colIndexAttachmentPath = colIndexCostCenter + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderGAAllowancePT = firstRow.GetCell(colIndexGAAllowancePT);
				ICell cellHeaderFinancing = firstRow.GetCell(colIndexFinancing);
				ICell cellHeaderMGI = firstRow.GetCell(colIndexMGI);
				ICell cellHeaderTaxCalculation = firstRow.GetCell(colIndexTaxCalculation);
				ICell cellHeaderFirstYearCR = firstRow.GetCell(colIndexFirstYearCR);
				ICell cellHeaderRenewalYearCR = firstRow.GetCell(colIndexRenewalYearCR);
				ICell cellHeaderReportRecOverride = firstRow.GetCell(colIndexReportRecOverride);
				ICell cellHeaderADGenDetail = firstRow.GetCell(colIndexADGenDetail);
				ICell cellHeaderProducerBonusReport = firstRow.GetCell(colIndexProducerBonusReport);
				ICell cellHeaderAALBonusReport = firstRow.GetCell(colIndexAALBonusReport);
				ICell cellHeaderSummaryOverride = firstRow.GetCell(colIndexSummaryOverride);
				ICell cellHeaderGAAllowancePersonal = firstRow.GetCell(colIndexGAAllowancePersonal);
				ICell cellHeaderSummaryAgentPaid = firstRow.GetCell(colIndexSummaryAgentPaid);
				ICell cellHeaderCFGManualPayment = firstRow.GetCell(colIndexCFGManualPayment);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderGAAllowancePT, workSheetName, colIndexGAAllowancePT + 1, "GA Allowance PT Path");
				this.ValidateCellHeaderName(cellHeaderFinancing, workSheetName, colIndexFinancing + 1, "Financing Path");
				this.ValidateCellHeaderName(cellHeaderMGI, workSheetName, colIndexMGI + 1, "MGI Path");
				this.ValidateCellHeaderName(cellHeaderTaxCalculation, workSheetName, colIndexTaxCalculation + 1, "Tax Calculation Path");
				this.ValidateCellHeaderName(cellHeaderFirstYearCR, workSheetName, colIndexFirstYearCR + 1, "First Year CR Path");
				this.ValidateCellHeaderName(cellHeaderRenewalYearCR, workSheetName, colIndexRenewalYearCR + 1, "Renewal Year CR Path");
				this.ValidateCellHeaderName(cellHeaderReportRecOverride, workSheetName, colIndexReportRecOverride + 1, "Report Rec Override Path");
				this.ValidateCellHeaderName(cellHeaderADGenDetail, workSheetName, colIndexADGenDetail + 1, "AD Gen Detail Path");
				this.ValidateCellHeaderName(cellHeaderProducerBonusReport, workSheetName, colIndexProducerBonusReport + 1, "Producer Bonus Report Path");
				this.ValidateCellHeaderName(cellHeaderAALBonusReport, workSheetName, colIndexAALBonusReport + 1, "AAL Bonus Report Path");
				this.ValidateCellHeaderName(cellHeaderSummaryOverride, workSheetName, colIndexSummaryOverride + 1, "Summary Override Path");
				this.ValidateCellHeaderName(cellHeaderGAAllowancePersonal, workSheetName, colIndexGAAllowancePersonal + 1, "GA Allowance Personal Path");
				this.ValidateCellHeaderName(cellHeaderSummaryAgentPaid, workSheetName, colIndexSummaryAgentPaid + 1, "Summary Agent Paid Path");
				this.ValidateCellHeaderName(cellHeaderCFGManualPayment, workSheetName, colIndexCFGManualPayment + 1, "CFG Manual Payment Path");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center Name");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVAgencyCommissionModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							GAAllowancePTPath = this.TryGetString(currentRow.GetCell(colIndexGAAllowancePT), workSheetName, cellHeaderGAAllowancePT.StringCellValue, rowNum),
							FinancingPath = this.TryGetString(currentRow.GetCell(colIndexFinancing), workSheetName, cellHeaderFinancing.StringCellValue, rowNum),
							MGIPath = this.TryGetString(currentRow.GetCell(colIndexMGI), workSheetName, cellHeaderMGI.StringCellValue, rowNum),
							TaxCalculationPath = this.TryGetString(currentRow.GetCell(colIndexTaxCalculation), workSheetName, cellHeaderTaxCalculation.StringCellValue, rowNum),
							FirstYearCRPath = this.TryGetString(currentRow.GetCell(colIndexFirstYearCR), workSheetName, cellHeaderFirstYearCR.StringCellValue, rowNum),
							RenewalYearCRPath = this.TryGetString(currentRow.GetCell(colIndexRenewalYearCR), workSheetName, cellHeaderRenewalYearCR.StringCellValue, rowNum),
							ReportRecOverridePath = this.TryGetString(currentRow.GetCell(colIndexReportRecOverride), workSheetName, cellHeaderReportRecOverride.StringCellValue, rowNum),
							ADGenDetailPath = this.TryGetString(currentRow.GetCell(colIndexADGenDetail), workSheetName, cellHeaderADGenDetail.StringCellValue, rowNum),
							ProducerBonusReportPath = this.TryGetString(currentRow.GetCell(colIndexProducerBonusReport), workSheetName, cellHeaderProducerBonusReport.StringCellValue, rowNum),
							AALBonusReportPath = this.TryGetString(currentRow.GetCell(colIndexAALBonusReport), workSheetName, cellHeaderAALBonusReport.StringCellValue, rowNum),
							SummaryOverridePath = this.TryGetString(currentRow.GetCell(colIndexSummaryOverride), workSheetName, cellHeaderSummaryOverride.StringCellValue, rowNum),
							GAAllowancePersonalPath = this.TryGetString(currentRow.GetCell(colIndexGAAllowancePersonal), workSheetName, cellHeaderGAAllowancePersonal.StringCellValue, rowNum),
							SummaryAgentPaidPath = this.TryGetString(currentRow.GetCell(colIndexSummaryAgentPaid), workSheetName, cellHeaderSummaryAgentPaid.StringCellValue, rowNum),
							CFGManualPaymentPath = this.TryGetString(currentRow.GetCell(colIndexCFGManualPayment), workSheetName, cellHeaderCFGManualPayment.StringCellValue, rowNum),
							CostCenterName = this.TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVAgencyCommission} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVSalesForceModels> ReadPVSalesForce(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVSalesForceModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVSalesForce).ToList();
			try
			{
				string workSheetName = SheetConstant.PVSalesForce;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexCostCenterName = colIndexPaymentPurpose + 1;
				int colIndexBusinessName = colIndexCostCenterName + 1;
				int colIndexPayment = colIndexBusinessName + 1;
				int colIndexShortDesc = colIndexPayment + 1;
				int colIndexFilePath = colIndexShortDesc + 1;
				int colIndexIsProposal = colIndexFilePath + 1;
				int colIndexProposalPath = colIndexIsProposal + 1;
				int colIndexBeneficiaryType = colIndexProposalPath + 1;
				int colIndexPeriodOfPayment = colIndexBeneficiaryType + 1;
				int colIndexAttachmentPath = colIndexPeriodOfPayment + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderCostCenterName = firstRow.GetCell(colIndexCostCenterName);
				ICell cellHeaderBusinessName = firstRow.GetCell(colIndexBusinessName);
				ICell cellHeaderPayment = firstRow.GetCell(colIndexPayment);
				ICell cellHeaderShortDesc = firstRow.GetCell(colIndexShortDesc);
				ICell cellHeaderFilePath = firstRow.GetCell(colIndexFilePath);
				ICell cellHeaderIsProposal = firstRow.GetCell(colIndexIsProposal);
				ICell cellHeaderProposalPath = firstRow.GetCell(colIndexProposalPath);
				ICell cellHeaderBeneficiaryType = firstRow.GetCell(colIndexBeneficiaryType);
				ICell cellHeaderPeriodOfPayment = firstRow.GetCell(colIndexPeriodOfPayment);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderCostCenterName, workSheetName, colIndexCostCenterName + 1, "Cost Center Name");
				this.ValidateCellHeaderName(cellHeaderBusinessName, workSheetName, colIndexBusinessName + 1, "Business Name");
				this.ValidateCellHeaderName(cellHeaderPayment, workSheetName, colIndexPayment + 1, "Payment");
				this.ValidateCellHeaderName(cellHeaderShortDesc, workSheetName, colIndexShortDesc + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderFilePath, workSheetName, colIndexFilePath + 1, "File Path");
				this.ValidateCellHeaderName(cellHeaderIsProposal, workSheetName, colIndexIsProposal + 1, "With PO / Proposal or Not");
				this.ValidateCellHeaderName(cellHeaderProposalPath, workSheetName, colIndexProposalPath + 1, "Proposal Path");
				this.ValidateCellHeaderName(cellHeaderBeneficiaryType, workSheetName, colIndexBeneficiaryType + 1, "Beneficiary Customer Type");
				this.ValidateCellHeaderName(cellHeaderPeriodOfPayment, workSheetName, colIndexPeriodOfPayment + 1, "Period of Payment");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVSalesForceModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							CostCenterName = this.TryGetString(currentRow.GetCell(colIndexCostCenterName), workSheetName, cellHeaderCostCenterName.StringCellValue, rowNum),
							BusinessName = this.TryGetString(currentRow.GetCell(colIndexBusinessName), workSheetName, cellHeaderBusinessName.StringCellValue, rowNum),
							Payment = this.TryGetString(currentRow.GetCell(colIndexPayment), workSheetName, cellHeaderPayment.StringCellValue, rowNum),
							ShortDescription = this.TryGetString(currentRow.GetCell(colIndexShortDesc), workSheetName, cellHeaderShortDesc.StringCellValue, rowNum),
							FilePath = this.TryGetString(currentRow.GetCell(colIndexFilePath), workSheetName, cellHeaderFilePath.StringCellValue, rowNum),
							IsProposal = this.TryGetBoolean(currentRow.GetCell(colIndexIsProposal), workSheetName, cellHeaderIsProposal.StringCellValue, rowNum),
							ProposalPath = this.TryGetString(currentRow.GetCell(colIndexProposalPath), workSheetName, cellHeaderProposalPath.StringCellValue, rowNum),
							BeneficiaryCustomerType = this.TryGetString(currentRow.GetCell(colIndexBeneficiaryType), workSheetName, cellHeaderBeneficiaryType.StringCellValue, rowNum),
							PeriodOfPayment = this.TryGetString(currentRow.GetCell(colIndexPeriodOfPayment), workSheetName, cellHeaderPeriodOfPayment.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVSalesForce} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVSalesForceModels> ReadPVRDD(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVSalesForceModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVRDD).ToList();
			try
			{
				string workSheetName = SheetConstant.PVRDD;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexCostCenterName = colIndexPaymentPurpose + 1;
				int colIndexBusinessName = colIndexCostCenterName + 1;
				int colIndexPayment = colIndexBusinessName + 1;
				int colIndexShortDesc = colIndexPayment + 1;
				int colIndexFilePath = colIndexShortDesc + 1;
				int colIndexIsProposal = colIndexFilePath + 1;
				int colIndexProposalPath = colIndexIsProposal + 1;
				int colIndexPeriodOfPayment = colIndexProposalPath + 1;
				int colIndexAttachmentPath = colIndexPeriodOfPayment + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderCostCenterName = firstRow.GetCell(colIndexCostCenterName);
				ICell cellHeaderBusinessName = firstRow.GetCell(colIndexBusinessName);
				ICell cellHeaderPayment = firstRow.GetCell(colIndexPayment);
				ICell cellHeaderShortDesc = firstRow.GetCell(colIndexShortDesc);
				ICell cellHeaderFilePath = firstRow.GetCell(colIndexFilePath);
				ICell cellHeaderIsProposal = firstRow.GetCell(colIndexIsProposal);
				ICell cellHeaderProposalPath = firstRow.GetCell(colIndexProposalPath);
				ICell cellHeaderPeriodOfPayment = firstRow.GetCell(colIndexPeriodOfPayment);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderCostCenterName, workSheetName, colIndexCostCenterName + 1, "Cost Center Name");
				this.ValidateCellHeaderName(cellHeaderBusinessName, workSheetName, colIndexBusinessName + 1, "Business Name");
				this.ValidateCellHeaderName(cellHeaderPayment, workSheetName, colIndexPayment + 1, "Payment");
				this.ValidateCellHeaderName(cellHeaderShortDesc, workSheetName, colIndexShortDesc + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderFilePath, workSheetName, colIndexFilePath + 1, "File Path");
				this.ValidateCellHeaderName(cellHeaderIsProposal, workSheetName, colIndexIsProposal + 1, "With PO / Proposal or Not");
				this.ValidateCellHeaderName(cellHeaderProposalPath, workSheetName, colIndexProposalPath + 1, "Proposal Path");
				this.ValidateCellHeaderName(cellHeaderPeriodOfPayment, workSheetName, colIndexPeriodOfPayment + 1, "Period of Payment");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVSalesForceModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							CostCenterName = this.TryGetString(currentRow.GetCell(colIndexCostCenterName), workSheetName, cellHeaderCostCenterName.StringCellValue, rowNum),
							BusinessName = this.TryGetString(currentRow.GetCell(colIndexBusinessName), workSheetName, cellHeaderBusinessName.StringCellValue, rowNum),
							Payment = this.TryGetString(currentRow.GetCell(colIndexPayment), workSheetName, cellHeaderPayment.StringCellValue, rowNum),
							ShortDescription = this.TryGetString(currentRow.GetCell(colIndexShortDesc), workSheetName, cellHeaderShortDesc.StringCellValue, rowNum),
							FilePath = this.TryGetString(currentRow.GetCell(colIndexFilePath), workSheetName, cellHeaderFilePath.StringCellValue, rowNum),
							IsProposal = this.TryGetBoolean(currentRow.GetCell(colIndexIsProposal), workSheetName, cellHeaderIsProposal.StringCellValue, rowNum),
							ProposalPath = this.TryGetString(currentRow.GetCell(colIndexProposalPath), workSheetName, cellHeaderProposalPath.StringCellValue, rowNum),
							PeriodOfPayment = this.TryGetString(currentRow.GetCell(colIndexPeriodOfPayment), workSheetName, cellHeaderPeriodOfPayment.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVRDD} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVVendorModels> ReadPVVendor(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVVendorModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVVendor).ToList();
			try
			{
				string workSheetName = SheetConstant.PVVendor;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexPayee = colIndexPaymentPurpose + 1;
				int colIndexIsRecurring = colIndexPayee + 1;
				int colIndexPO = colIndexIsRecurring + 1;
				int colIndexGR = colIndexPO + 1;
				int colIndexAttachmentPath = colIndexGR + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				int colIndexRequestorName = colIndexRemarks + 1;
				int colIndexCostCenter = colIndexRequestorName + 1;
				int colIndexAccrualCode = colIndexCostCenter + 1;
				int colIndexBusinessName = colIndexAccrualCode + 1;
				int colIndexLongDesc = colIndexBusinessName + 1;
				int colIndexShortDesc = colIndexLongDesc + 1;
				int colIndexIsProposal = colIndexShortDesc + 1;
				int colIndexProposalPath = colIndexIsProposal + 1;
				int colIndexInvoiceNo = colIndexProposalPath + 1;
				int colIndexInvoiceDate = colIndexInvoiceNo + 1;
				int colIndexInvoiceDueDate = colIndexInvoiceDate + 1;
				int colIndexInvoicePath = colIndexInvoiceDueDate + 1;
				int colIndexBankAccount = colIndexInvoicePath + 1;
				int colIndexGrossAmount = colIndexBankAccount + 1;
				int colIndexDetailRemarks = colIndexGrossAmount + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderPayee = firstRow.GetCell(colIndexPayee);
				ICell cellHeaderIsRecurring = firstRow.GetCell(colIndexIsRecurring);
				ICell cellHeaderPO = firstRow.GetCell(colIndexPO);
				ICell cellHeaderGR = firstRow.GetCell(colIndexGR);

				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderRequestorName = firstRow.GetCell(colIndexRequestorName);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderAccrualCode = firstRow.GetCell(colIndexAccrualCode);
				ICell cellHeaderBusinessName = firstRow.GetCell(colIndexBusinessName);
				ICell cellHeaderLongDesc = firstRow.GetCell(colIndexLongDesc);
				ICell cellHeaderShortDesc = firstRow.GetCell(colIndexShortDesc);
				ICell cellHeaderIsProposal = firstRow.GetCell(colIndexIsProposal);
				ICell cellHeaderProposalPath = firstRow.GetCell(colIndexProposalPath);
				ICell cellHeaderInvoiceNo = firstRow.GetCell(colIndexInvoiceNo);
				ICell cellHeaderInvoiceDate = firstRow.GetCell(colIndexInvoiceDate);
				ICell cellHeaderInvoiceDueDate = firstRow.GetCell(colIndexInvoiceDueDate);
				ICell cellHeaderInvoicePath = firstRow.GetCell(colIndexInvoicePath);
				ICell cellHeaderBankAccount = firstRow.GetCell(colIndexBankAccount);
				ICell cellHeaderGrossAmount = firstRow.GetCell(colIndexGrossAmount);
				ICell cellHeaderDetailRemarks = firstRow.GetCell(colIndexDetailRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderPayee, workSheetName, colIndexPayee + 1, "Payee");
				this.ValidateCellHeaderName(cellHeaderIsRecurring, workSheetName, colIndexIsRecurring + 1, "Is Recurring");
				this.ValidateCellHeaderName(cellHeaderPO, workSheetName, colIndexPO + 1, "PO");
				this.ValidateCellHeaderName(cellHeaderGR, workSheetName, colIndexGR + 1, "GR");

				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderRequestorName, workSheetName, colIndexRequestorName + 1, "Requestor Name");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center");
				this.ValidateCellHeaderName(cellHeaderAccrualCode, workSheetName, colIndexAccrualCode + 1, "Accrual Code");
				this.ValidateCellHeaderName(cellHeaderBusinessName, workSheetName, colIndexBusinessName + 1, "Business Name");
				this.ValidateCellHeaderName(cellHeaderLongDesc, workSheetName, colIndexLongDesc + 1, "Long Description");
				this.ValidateCellHeaderName(cellHeaderShortDesc, workSheetName, colIndexShortDesc + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderIsProposal, workSheetName, colIndexIsProposal + 1, "Is Proposal");
				this.ValidateCellHeaderName(cellHeaderProposalPath, workSheetName, colIndexProposalPath + 1, "Proposal Path");
				this.ValidateCellHeaderName(cellHeaderInvoiceNo, workSheetName, colIndexInvoiceNo + 1, "Invoice No");
				this.ValidateCellHeaderName(cellHeaderInvoiceDate, workSheetName, colIndexInvoiceDate + 1, "Invoice Date");
				this.ValidateCellHeaderName(cellHeaderInvoiceDueDate, workSheetName, colIndexInvoiceDueDate + 1, "Invoice Due Date");
				this.ValidateCellHeaderName(cellHeaderInvoicePath, workSheetName, colIndexInvoicePath + 1, "Invoice Path");
				this.ValidateCellHeaderName(cellHeaderBankAccount, workSheetName, colIndexBankAccount + 1, "Bank Account");
				this.ValidateCellHeaderName(cellHeaderGrossAmount, workSheetName, colIndexGrossAmount + 1, "Gross Amount");
				this.ValidateCellHeaderName(cellHeaderDetailRemarks, workSheetName, colIndexDetailRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexDetailRemarks))
					{
						break;
					}

					#region Assign Param
					string idStr = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(idStr))
					{
						PVVendorModels param = new();
						int id = int.Parse(idStr);
						if (plans.Any(x => x.TestCaseId.Equals(id)))
						{
							PVDetailTransactionLongModels detail = new()
							{
								RequestorName = this.TryGetString(currentRow.GetCell(colIndexRequestorName), workSheetName, cellHeaderRequestorName.StringCellValue, rowNum),
								CostCenter = this.TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
								AccrualCode = this.TryGetString(currentRow.GetCell(colIndexAccrualCode), workSheetName, cellHeaderAccrualCode.StringCellValue, rowNum),
								BusinessName = this.TryGetString(currentRow.GetCell(colIndexBusinessName), workSheetName, cellHeaderBusinessName.StringCellValue, rowNum),
								LongDescription = this.TryGetString(currentRow.GetCell(colIndexLongDesc), workSheetName, cellHeaderLongDesc.StringCellValue, rowNum),
								ShortDescription = this.TryGetString(currentRow.GetCell(colIndexShortDesc), workSheetName, cellHeaderShortDesc.StringCellValue, rowNum),
								IsProposal = this.TryGetBoolean(currentRow.GetCell(colIndexIsProposal), workSheetName, cellHeaderIsProposal.StringCellValue, rowNum),
								ProposalPath = this.TryGetString(currentRow.GetCell(colIndexProposalPath), workSheetName, cellHeaderProposalPath.StringCellValue, rowNum),
								InvoiceNo = this.TryGetString(currentRow.GetCell(colIndexInvoiceNo), workSheetName, cellHeaderInvoiceNo.StringCellValue, rowNum),
								InvoiceDate = this.TryGetString(currentRow.GetCell(colIndexInvoiceDate), workSheetName, cellHeaderInvoiceDate.StringCellValue, rowNum),
								InvoiceDueDate = this.TryGetString(currentRow.GetCell(colIndexInvoiceDueDate), workSheetName, cellHeaderInvoiceDueDate.StringCellValue, rowNum),
								InvoiceAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexInvoicePath), workSheetName, cellHeaderInvoicePath.StringCellValue, rowNum),
								BankAccountNo = this.TryGetString(currentRow.GetCell(colIndexBankAccount), workSheetName, cellHeaderBankAccount.StringCellValue, rowNum),
								GrossAmount = this.TryGetString(currentRow.GetCell(colIndexGrossAmount), workSheetName, cellHeaderGrossAmount.StringCellValue, rowNum),
								Remarks = this.TryGetString(currentRow.GetCell(colIndexDetailRemarks), workSheetName, cellHeaderDetailRemarks.StringCellValue, rowNum),
							};

							if (this._dictPVVendor.TryGetValue(id, out PVVendorModels? value))
							{
								value.AddDetails.Add(detail);
							}
							else
							{
								param.Row = rowNum;
								param.TestCaseId = id;
								param.DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum);
								param.Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1;
								param.Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum);
								param.PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum);
								param.Payee = this.TryGetString(currentRow.GetCell(colIndexPayee), workSheetName, cellHeaderPayee.StringCellValue, rowNum);
								param.IsRecurring = this.TryGetBoolean(currentRow.GetCell(colIndexIsRecurring), workSheetName, cellHeaderIsRecurring.StringCellValue, rowNum);
								param.PO = this.TryGetString(currentRow.GetCell(colIndexPO), workSheetName, cellHeaderPO.StringCellValue, rowNum);
								param.GR = this.TryGetString(currentRow.GetCell(colIndexGR), workSheetName, cellHeaderGR.StringCellValue, rowNum);
								param.AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum);
								param.AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum);
								param.Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum);

								param.AddDetails.Add(detail);
								this._dictPVVendor.Add(id, param);
							}

						}
					}
					#endregion
				}

				result.AddRange(this._dictPVVendor.Values);
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVVendor} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVSettlementModels> ReadPVSettlement(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVSettlementModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVSettlement).ToList();
			try
			{
				string workSheetName = SheetConstant.PVSettlement;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexRequestorName = colIndexPaymentPurpose + 1;
				int colIndexPaymentVoucherNo = colIndexRequestorName + 1;
				int colIndexAttachmentPath = colIndexPaymentVoucherNo + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;
				int colIndexBusinessName = colIndexRemarks + 1;
				int colIndexCostCenter = colIndexBusinessName + 1;
				int colIndexLongDescription = colIndexCostCenter + 1;
				int colIndexShortDescription = colIndexLongDescription + 1;
				int colIndexDate = colIndexShortDescription + 1;
				int colIndexAmount = colIndexDate + 1;
				int colIndexDetailAttachmentPath = colIndexAmount + 1;
				int colIndexDDate = colIndexDetailAttachmentPath + 1;
				int colIndexDDescription = colIndexDDate + 1;
				int colIndexDAmount = colIndexDDescription + 1;
				int colIndexDCurrency = colIndexDAmount + 1;
				int colIndexDRate = colIndexDCurrency + 1;
				int colIndexDAttachmentPath = colIndexDRate + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderRequestorName = firstRow.GetCell(colIndexRequestorName);
				ICell cellHeaderPaymentVoucherNo = firstRow.GetCell(colIndexPaymentVoucherNo);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderBusinessName = firstRow.GetCell(colIndexBusinessName);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderLongDescription = firstRow.GetCell(colIndexLongDescription);
				ICell cellHeaderShortDescription = firstRow.GetCell(colIndexShortDescription);
				ICell cellHeaderDate = firstRow.GetCell(colIndexDate);
				ICell cellHeaderAmount = firstRow.GetCell(colIndexAmount);
				ICell cellHeaderDetailAttachmentPath = firstRow.GetCell(colIndexDetailAttachmentPath);
				ICell cellHeaderDDate = firstRow.GetCell(colIndexDDate);
				ICell cellHeaderDDesc = firstRow.GetCell(colIndexDDescription);
				ICell cellHeaderDAmount = firstRow.GetCell(colIndexDAmount);
				ICell cellHeaderDCurrency = firstRow.GetCell(colIndexDCurrency);
				ICell cellHeaderDRate = firstRow.GetCell(colIndexDRate);
				ICell cellHeaderDAttachmentPath = firstRow.GetCell(colIndexDAttachmentPath);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderRequestorName, workSheetName, colIndexRequestorName + 1, "Requestor Name");
				this.ValidateCellHeaderName(cellHeaderPaymentVoucherNo, workSheetName, colIndexPaymentVoucherNo + 1, "Payment Voucher No");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderBusinessName, workSheetName, colIndexBusinessName + 1, "Business Name");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center");
				this.ValidateCellHeaderName(cellHeaderLongDescription, workSheetName, colIndexLongDescription + 1, "Long Description");
				this.ValidateCellHeaderName(cellHeaderShortDescription, workSheetName, colIndexShortDescription + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderDate, workSheetName, colIndexDate + 1, "Date");
				this.ValidateCellHeaderName(cellHeaderAmount, workSheetName, colIndexAmount + 1, "Amount");
				this.ValidateCellHeaderName(cellHeaderDetailAttachmentPath, workSheetName, colIndexDetailAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderDDate, workSheetName, colIndexDDate + 1, "Date");
				this.ValidateCellHeaderName(cellHeaderDDesc, workSheetName, colIndexDDescription + 1, "Description");
				this.ValidateCellHeaderName(cellHeaderDAmount, workSheetName, colIndexDAmount + 1, "Amount");
				this.ValidateCellHeaderName(cellHeaderDCurrency, workSheetName, colIndexDCurrency + 1, "Currency");
				this.ValidateCellHeaderName(cellHeaderDRate, workSheetName, colIndexDRate + 1, "Rate");
				this.ValidateCellHeaderName(cellHeaderDAttachmentPath, workSheetName, colIndexDAttachmentPath + 1, "Attachment Path");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexDAttachmentPath))
					{
						break;
					}

					#region Assign Param
					string idStr = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(idStr))
					{
						PVSettlementModels param = new();
						int id = int.Parse(idStr);
						if (plans.Any(x => x.TestCaseId.Equals(id)))
						{
							PVDetailTransactionModels detail = new()
							{
								BusinessName = this.TryGetString(currentRow.GetCell(colIndexBusinessName), workSheetName, cellHeaderBusinessName.StringCellValue, rowNum),
								CostCenter = this.TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
								LongDescription = this.TryGetString(currentRow.GetCell(colIndexLongDescription), workSheetName, cellHeaderLongDescription.StringCellValue, rowNum),
								ShortDescription = this.TryGetString(currentRow.GetCell(colIndexShortDescription), workSheetName, cellHeaderShortDescription.StringCellValue, rowNum),
								Date = this.TryGetString(currentRow.GetCell(colIndexDate), workSheetName, cellHeaderDate.StringCellValue, rowNum),
								Amount = this.TryGetString(currentRow.GetCell(colIndexAmount), workSheetName, cellHeaderAmount.StringCellValue, rowNum),
								AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDetailAttachmentPath), workSheetName, cellHeaderDetailAttachmentPath.StringCellValue, rowNum)
							};

							PVDetailDetailTransactionModels detailDetail = new()
							{
								Date = this.TryGetString(currentRow.GetCell(colIndexDDate), workSheetName, cellHeaderDDate.StringCellValue, rowNum),
								Description = this.TryGetString(currentRow.GetCell(colIndexDDescription), workSheetName, cellHeaderDDesc.StringCellValue, rowNum),
								Amount = this.TryGetString(currentRow.GetCell(colIndexDAmount), workSheetName, cellHeaderDAmount.StringCellValue, rowNum),
								Currency = this.TryGetString(currentRow.GetCell(colIndexDCurrency), workSheetName, cellHeaderDCurrency.StringCellValue, rowNum),
								Rate = this.TryGetString(currentRow.GetCell(colIndexDRate), workSheetName, cellHeaderDRate.StringCellValue, rowNum),
								AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDAttachmentPath), workSheetName, cellHeaderDAttachmentPath.StringCellValue, rowNum)
							};

							if (this._dictPVSettlement.TryGetValue(id, out PVSettlementModels? value))
							{
								if (AreAllPropertiesSet(detailDetail))
									detail.Details.Add(detailDetail);

								value.AddDetails.Add(detail);
							}
							else
							{
								param.Row = rowNum;
								param.TestCaseId = id;
								param.DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum);
								param.Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1;
								param.Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum);
								param.PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum);
								param.RequestorName = this.TryGetString(currentRow.GetCell(colIndexRequestorName), workSheetName, cellHeaderRequestorName.StringCellValue, rowNum);
								param.PaymentVoucherNo = this.TryGetString(currentRow.GetCell(colIndexPaymentVoucherNo), workSheetName, cellHeaderPaymentVoucherNo.StringCellValue, rowNum);
								param.AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum);
								param.AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum);
								param.Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum);

								if (AreAllPropertiesSet(detailDetail))
									detail.Details.Add(detailDetail);
								param.AddDetails.Add(detail);
								this._dictPVSettlement.Add(id, param);
							}
						}

					}
					#endregion
				}
				result.AddRange(this._dictPVSettlement.Values);
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVSettlement} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVAdvanceModels> ReadPVAdvance(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVAdvanceModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVAdvance).ToList();
			try
			{
				string workSheetName = SheetConstant.PVAdvance;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexRequestorName = colIndexPaymentPurpose + 1;
				int colIndexBankName = colIndexRequestorName + 1;
				int colIndexBankAccountName = colIndexBankName + 1;
				int colIndexBankAccountNo = colIndexBankAccountName + 1;
				int colIndexStartDate = colIndexBankAccountNo + 1;
				int colIndexEndDate = colIndexStartDate + 1;
				int colIndexAmount = colIndexEndDate + 1;
				int colIndexReason = colIndexAmount + 1;
				int colIndexAttachmentPath = colIndexReason + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderRequestorName = firstRow.GetCell(colIndexRequestorName);
				ICell cellHeaderBankName = firstRow.GetCell(colIndexBankName);
				ICell cellHeaderBankAccountName = firstRow.GetCell(colIndexBankAccountName);
				ICell cellHeaderBankAccountNo = firstRow.GetCell(colIndexBankAccountNo);
				ICell cellHeaderStartDate = firstRow.GetCell(colIndexStartDate);
				ICell cellHeaderEndDate = firstRow.GetCell(colIndexEndDate);
				ICell cellHeaderAmount = firstRow.GetCell(colIndexAmount);
				ICell cellHeaderReason = firstRow.GetCell(colIndexReason);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderRequestorName, workSheetName, colIndexRequestorName + 1, "Requestor Name");
				this.ValidateCellHeaderName(cellHeaderBankName, workSheetName, colIndexBankName + 1, "Bank Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountName, workSheetName, colIndexBankAccountName + 1, "Requestor Bank Account Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountNo, workSheetName, colIndexBankAccountNo + 1, "Requestor Bank Account No");
				this.ValidateCellHeaderName(cellHeaderStartDate, workSheetName, colIndexStartDate + 1, "Start Date");
				this.ValidateCellHeaderName(cellHeaderEndDate, workSheetName, colIndexEndDate + 1, "End Date");
				this.ValidateCellHeaderName(cellHeaderAmount, workSheetName, colIndexAmount + 1, "Amount");
				this.ValidateCellHeaderName(cellHeaderReason, workSheetName, colIndexReason + 1, "Reason");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						PVAdvanceModels param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum),
							PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum),
							RequestorName = this.TryGetString(currentRow.GetCell(colIndexRequestorName), workSheetName, cellHeaderRequestorName.StringCellValue, rowNum),
							BankName = this.TryGetString(currentRow.GetCell(colIndexBankName), workSheetName, cellHeaderBankName.StringCellValue, rowNum),
							RequestorBankAccountName = this.TryGetString(currentRow.GetCell(colIndexBankAccountName), workSheetName, cellHeaderBankAccountName.StringCellValue, rowNum),
							RequestorBankAccountNo = this.TryGetString(currentRow.GetCell(colIndexBankAccountNo), workSheetName, cellHeaderBankAccountNo.StringCellValue, rowNum),
							StartDate = this.TryGetString(currentRow.GetCell(colIndexStartDate), workSheetName, cellHeaderStartDate.StringCellValue, rowNum),
							EndDate = this.TryGetString(currentRow.GetCell(colIndexEndDate), workSheetName, cellHeaderEndDate.StringCellValue, rowNum),
							Amount = this.TryGetString(currentRow.GetCell(colIndexAmount), workSheetName, cellHeaderAmount.StringCellValue, rowNum),
							Reason = this.TryGetString(currentRow.GetCell(colIndexReason), workSheetName, cellHeaderReason.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum)
						};

						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVAdvance} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<PVReimbursementModels> ReadPVReimbursement(IWorkbook wb, List<TestPlan> plans)
		{
			List<PVReimbursementModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PVReimbursement).ToList();
			try
			{
				string workSheetName = SheetConstant.PVReimbursement;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexTitle = colIndexSeq + 1;
				int colIndexPaymentPurpose = colIndexTitle + 1;
				int colIndexRequestorName = colIndexPaymentPurpose + 1;
				int colIndexBankName = colIndexRequestorName + 1;
				int colIndexBankAccountName = colIndexBankName + 1;
				int colIndexBankAccountNo = colIndexBankAccountName + 1;
				int colIndexAttachmentPath = colIndexBankAccountNo + 1;
				int colIndexAttachmentDescription = colIndexAttachmentPath + 1;
				int colIndexRemarks = colIndexAttachmentDescription + 1;
				int colIndexBusinessName = colIndexRemarks + 1;
				int colIndexCostCenter = colIndexBusinessName + 1;
				int colIndexAccrualCode = colIndexCostCenter + 1;
				int colIndexLongDescription = colIndexAccrualCode + 1;
				int colIndexShortDescription = colIndexLongDescription + 1;
				int colIndexAmount = colIndexShortDescription + 1;
				int colIndexDetailAttachmentPath = colIndexAmount + 1;
				int colIndexIsProposal = colIndexDetailAttachmentPath + 1;
				int colIndexProposalPath = colIndexIsProposal + 1;
				int colIndexDDate = colIndexProposalPath + 1;
				int colIndexDDescription = colIndexDDate + 1;
				int colIndexDAmount = colIndexDDescription + 1;
				int colIndexDCurrency = colIndexDAmount + 1;
				int colIndexDRate = colIndexDCurrency + 1;
				int colIndexDAttachmentPath = colIndexDRate + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderTitle = firstRow.GetCell(colIndexTitle);
				ICell cellHeaderPaymentPurpose = firstRow.GetCell(colIndexPaymentPurpose);
				ICell cellHeaderRequestorName = firstRow.GetCell(colIndexRequestorName);
				ICell cellHeaderBankName = firstRow.GetCell(colIndexBankName);
				ICell cellHeaderBankAccountName = firstRow.GetCell(colIndexBankAccountName);
				ICell cellHeaderBankAccountNo = firstRow.GetCell(colIndexBankAccountNo);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderBusinessName = firstRow.GetCell(colIndexBusinessName);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderAccrualCode = firstRow.GetCell(colIndexAccrualCode);
				ICell cellHeaderLongDescription = firstRow.GetCell(colIndexLongDescription);
				ICell cellHeaderShortDescription = firstRow.GetCell(colIndexShortDescription);
				ICell cellHeaderAmount = firstRow.GetCell(colIndexAmount);
				ICell cellHeaderDetailAttachmentPath = firstRow.GetCell(colIndexDetailAttachmentPath);
				ICell cellHeaderIsProposal = firstRow.GetCell(colIndexIsProposal);
				ICell cellHeaderProposalPath = firstRow.GetCell(colIndexProposalPath);
				ICell cellHeaderDDate = firstRow.GetCell(colIndexDDate);
				ICell cellHeaderDDesc = firstRow.GetCell(colIndexDDescription);
				ICell cellHeaderDAmount = firstRow.GetCell(colIndexDAmount);
				ICell cellHeaderDCurrency = firstRow.GetCell(colIndexDCurrency);
				ICell cellHeaderDRate = firstRow.GetCell(colIndexDRate);
				ICell cellHeaderDAttachmentPath = firstRow.GetCell(colIndexDAttachmentPath);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderTitle, workSheetName, colIndexTitle + 1, "Title");
				this.ValidateCellHeaderName(cellHeaderPaymentPurpose, workSheetName, colIndexPaymentPurpose + 1, "Payment Purpose");
				this.ValidateCellHeaderName(cellHeaderRequestorName, workSheetName, colIndexRequestorName + 1, "Requestor Name");
				this.ValidateCellHeaderName(cellHeaderBankName, workSheetName, colIndexBankName + 1, "Bank Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountName, workSheetName, colIndexBankAccountName + 1, "Requestor Bank Account Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountNo, workSheetName, colIndexBankAccountNo + 1, "Requestor Bank Account No");
				this.ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderBusinessName, workSheetName, colIndexBusinessName + 1, "Business Name");
				this.ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center");
				this.ValidateCellHeaderName(cellHeaderAccrualCode, workSheetName, colIndexAccrualCode + 1, "Accrual Code");
				this.ValidateCellHeaderName(cellHeaderLongDescription, workSheetName, colIndexLongDescription + 1, "Long Description");
				this.ValidateCellHeaderName(cellHeaderShortDescription, workSheetName, colIndexShortDescription + 1, "Short Description");
				this.ValidateCellHeaderName(cellHeaderAmount, workSheetName, colIndexAmount + 1, "Amount");
				this.ValidateCellHeaderName(cellHeaderDetailAttachmentPath, workSheetName, colIndexDetailAttachmentPath + 1, "Detail Attachment Path");
				this.ValidateCellHeaderName(cellHeaderIsProposal, workSheetName, colIndexIsProposal + 1, "With PO / Proposal or Not");
				this.ValidateCellHeaderName(cellHeaderDDate, workSheetName, colIndexDDate + 1, "Date");
				this.ValidateCellHeaderName(cellHeaderDDesc, workSheetName, colIndexDDescription + 1, "Description");
				this.ValidateCellHeaderName(cellHeaderDAmount, workSheetName, colIndexDAmount + 1, "Amount");
				this.ValidateCellHeaderName(cellHeaderDCurrency, workSheetName, colIndexDCurrency + 1, "Currency");
				this.ValidateCellHeaderName(cellHeaderDRate, workSheetName, colIndexDRate + 1, "Rate");
				this.ValidateCellHeaderName(cellHeaderDAttachmentPath, workSheetName, colIndexDAttachmentPath + 1, "Attachment Path");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexDAttachmentPath))
					{
						break;
					}

					#region Assign Param
					string idStr = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(idStr))
					{
						PVReimbursementModels param = new();
						int id = int.Parse(idStr);
						if (plans.Any(x => x.TestCaseId.Equals(id)))
						{
							PVDetailTransactionModels detail = new()
							{
								BusinessName = this.TryGetString(currentRow.GetCell(colIndexBusinessName), workSheetName, cellHeaderBusinessName.StringCellValue, rowNum),
								CostCenter = this.TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum),
								AccrualCode = this.TryGetString(currentRow.GetCell(colIndexAccrualCode), workSheetName, cellHeaderAccrualCode.StringCellValue, rowNum),
								LongDescription = this.TryGetString(currentRow.GetCell(colIndexLongDescription), workSheetName, cellHeaderLongDescription.StringCellValue, rowNum),
								ShortDescription = this.TryGetString(currentRow.GetCell(colIndexShortDescription), workSheetName, cellHeaderShortDescription.StringCellValue, rowNum),
								Amount = this.TryGetString(currentRow.GetCell(colIndexAmount), workSheetName, cellHeaderAmount.StringCellValue, rowNum),
								AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDetailAttachmentPath), workSheetName, cellHeaderDetailAttachmentPath.StringCellValue, rowNum),
								IsProposal = this.TryGetBoolean(currentRow.GetCell(colIndexIsProposal), workSheetName, cellHeaderIsProposal.StringCellValue, rowNum),
								ProposalPath = this.TryGetString(currentRow.GetCell(colIndexProposalPath), workSheetName, cellHeaderProposalPath.StringCellValue, rowNum)
							};

							PVDetailDetailTransactionModels detailDetail = new()
							{
								Date = this.TryGetString(currentRow.GetCell(colIndexDDate), workSheetName, cellHeaderDDate.StringCellValue, rowNum),
								Description = this.TryGetString(currentRow.GetCell(colIndexDDescription), workSheetName, cellHeaderDDesc.StringCellValue, rowNum),
								Amount = this.TryGetString(currentRow.GetCell(colIndexDAmount), workSheetName, cellHeaderDAmount.StringCellValue, rowNum),
								Currency = this.TryGetString(currentRow.GetCell(colIndexDCurrency), workSheetName, cellHeaderDCurrency.StringCellValue, rowNum),
								Rate = this.TryGetString(currentRow.GetCell(colIndexDRate), workSheetName, cellHeaderDRate.StringCellValue, rowNum),
								AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexDAttachmentPath), workSheetName, cellHeaderDAttachmentPath.StringCellValue, rowNum)
							};

							if (this._dictPVReimbursement.TryGetValue(id, out PVReimbursementModels? value))
							{
								if (AreAllPropertiesSet(detailDetail))
									detail.Details.Add(detailDetail);

								value.AddDetails.Add(detail);
							}
							else
							{
								param.Row = rowNum;
								param.TestCaseId = id;
								param.DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum);
								param.Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1;
								param.Title = this.TryGetString(currentRow.GetCell(colIndexTitle), workSheetName, cellHeaderTitle.StringCellValue, rowNum);
								param.PaymentPurpose = this.TryGetString(currentRow.GetCell(colIndexPaymentPurpose), workSheetName, cellHeaderPaymentPurpose.StringCellValue, rowNum);
								param.RequestorName = this.TryGetString(currentRow.GetCell(colIndexRequestorName), workSheetName, cellHeaderRequestorName.StringCellValue, rowNum);
								param.BankName = this.TryGetString(currentRow.GetCell(colIndexBankName), workSheetName, cellHeaderBankName.StringCellValue, rowNum);
								param.RequestorBankAccountName = this.TryGetString(currentRow.GetCell(colIndexBankAccountName), workSheetName, cellHeaderBankAccountName.StringCellValue, rowNum);
								param.RequestorBankAccountNo = this.TryGetString(currentRow.GetCell(colIndexBankAccountNo), workSheetName, cellHeaderBankAccountNo.StringCellValue, rowNum);
								param.AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum);
								param.AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum);
								param.Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum);

								if (AreAllPropertiesSet(detailDetail))
									detail.Details.Add(detailDetail);
								param.AddDetails.Add(detail);
								this._dictPVReimbursement.Add(id, param);
							}
						}
					}
					#endregion
				}
				result.AddRange(this._dictPVReimbursement.Values);
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.PVReimbursement} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<VendorDatabase> ReadVendorDatabase(IWorkbook wb, List<TestPlan> plans)
		{
			List<VendorDatabase> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.VendorDatabase).ToList();
			try
			{
				string workSheetName = SheetConstant.VendorDatabase;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				#region Column Index
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexVendorName = colIndexSeq + 1;
				int colIndexVendorType = colIndexVendorName + 1;
				int colIndexCountry = colIndexVendorType + 1;
				int colIndexCity = colIndexCountry + 1;
				int colIndexAddress = colIndexCity + 1;
				int colIndexPhone1 = colIndexAddress + 1;
				int colIndexPhone2 = colIndexPhone1 + 1;
				int colIndexFax = colIndexPhone2 + 1;
				int colIndexEmail = colIndexFax + 1;
				int colIndexIsHaveCertificateofDomicile = colIndexEmail + 1;
				int colIndexCoDAttachmentPath = colIndexIsHaveCertificateofDomicile + 1;
				int colIndexCoDExpiryDate = colIndexCoDAttachmentPath + 1;
				int colIndexIsTaxableEntrepreneur = colIndexCoDExpiryDate + 1;
				int colIndexSPPKPPath = colIndexIsTaxableEntrepreneur + 1;
				int colIndexIsUnderWithholdingTax = colIndexSPPKPPath + 1;
				int colIndexRemarks = colIndexIsUnderWithholdingTax + 1;
				int colIndexAttacmentPath = colIndexRemarks + 1;
				int colIndexAttachmentDescription = colIndexAttacmentPath + 1;
				int colIndexIsHaveVendorTaxId = colIndexAttachmentDescription + 1;
				int colIndexTaxIdNo = colIndexIsHaveVendorTaxId + 1;
				int colIndexTaxIdName = colIndexTaxIdNo + 1;
				int colIndexTaxIdAddress = colIndexTaxIdName + 1;
				int colIndexTaxIdAttachmentPath = colIndexTaxIdAddress + 1;
				int colIndexTaxIdRemarks = colIndexTaxIdAttachmentPath + 1;
				#endregion

				#region Cell
				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderVendorName = firstRow.GetCell(colIndexVendorName);
				ICell cellHeaderVendorType = firstRow.GetCell(colIndexVendorType);
				ICell cellHeaderCountry = firstRow.GetCell(colIndexCountry);
				ICell cellHeaderCity = firstRow.GetCell(colIndexCity);
				ICell cellHeaderAddress = firstRow.GetCell(colIndexAddress);
				ICell cellHeaderPhone1 = firstRow.GetCell(colIndexPhone1);
				ICell cellHeaderPhone2 = firstRow.GetCell(colIndexPhone2);
				ICell cellHeaderFax = firstRow.GetCell(colIndexFax);
				ICell cellHeaderEmail = firstRow.GetCell(colIndexEmail);
				ICell cellHeaderIsHaveCertificateofDomicile = firstRow.GetCell(colIndexIsHaveCertificateofDomicile);
				ICell cellHeaderCoDAttachmentPath = firstRow.GetCell(colIndexCoDAttachmentPath);
				ICell cellHeaderCoDExpiryDate = firstRow.GetCell(colIndexCoDExpiryDate);
				ICell cellHeaderIsTaxableEntrepreneur = firstRow.GetCell(colIndexIsTaxableEntrepreneur);
				ICell cellHeaderSPPKPPath = firstRow.GetCell(colIndexSPPKPPath);
				ICell cellHeaderIsUnderWithholdingTax = firstRow.GetCell(colIndexIsUnderWithholdingTax);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderAttacmentPath = firstRow.GetCell(colIndexAttacmentPath);
				ICell cellHeaderAttachmentDescription = firstRow.GetCell(colIndexAttachmentDescription);
				ICell cellHeaderIsHaveVendorTaxId = firstRow.GetCell(colIndexIsHaveVendorTaxId);
				ICell cellHeaderTaxIdNo = firstRow.GetCell(colIndexTaxIdNo);
				ICell cellHeaderTaxIdName = firstRow.GetCell(colIndexTaxIdName);
				ICell cellHeaderTaxIdAddress = firstRow.GetCell(colIndexTaxIdAddress);
				ICell cellHeaderTaxIdAttachmentPath = firstRow.GetCell(colIndexTaxIdAttachmentPath);
				ICell cellHeaderTaxIdRemarks = firstRow.GetCell(colIndexTaxIdRemarks);
				#endregion

				#region Validate
				this.ValidateCellHeaderName(cellHeaderVendorName, workSheetName, colIndexVendorName + 1, "Vendor Name");
				this.ValidateCellHeaderName(cellHeaderVendorType, workSheetName, colIndexVendorType + 1, "Vendor Type");
				this.ValidateCellHeaderName(cellHeaderCountry, workSheetName, colIndexCountry + 1, "Country");
				this.ValidateCellHeaderName(cellHeaderCity, workSheetName, colIndexCity + 1, "City");
				this.ValidateCellHeaderName(cellHeaderAddress, workSheetName, colIndexAddress + 1, "Address");
				this.ValidateCellHeaderName(cellHeaderPhone1, workSheetName, colIndexPhone1 + 1, "Phone 1");
				this.ValidateCellHeaderName(cellHeaderPhone2, workSheetName, colIndexPhone2 + 1, "Phone 2");
				this.ValidateCellHeaderName(cellHeaderFax, workSheetName, colIndexFax + 1, "Fax");
				this.ValidateCellHeaderName(cellHeaderEmail, workSheetName, colIndexEmail + 1, "Email");
				this.ValidateCellHeaderName(cellHeaderIsHaveCertificateofDomicile, workSheetName, colIndexIsHaveCertificateofDomicile + 1, "Is Have Certificate of Domicile");
				this.ValidateCellHeaderName(cellHeaderCoDAttachmentPath, workSheetName, colIndexCoDAttachmentPath + 1, "CoD Attachment Path");
				this.ValidateCellHeaderName(cellHeaderCoDExpiryDate, workSheetName, colIndexCoDExpiryDate + 1, "CoD Expiry Date");
				this.ValidateCellHeaderName(cellHeaderIsTaxableEntrepreneur, workSheetName, colIndexIsTaxableEntrepreneur + 1, "Is Taxable Entrepreneur");
				this.ValidateCellHeaderName(cellHeaderSPPKPPath, workSheetName, colIndexSPPKPPath + 1, "SPPKP Path");
				this.ValidateCellHeaderName(cellHeaderIsUnderWithholdingTax, workSheetName, colIndexIsUnderWithholdingTax + 1, "Is Under Withholding Tax");
				this.ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				this.ValidateCellHeaderName(cellHeaderAttacmentPath, workSheetName, colIndexAttacmentPath + 1, "Attacment Path");
				this.ValidateCellHeaderName(cellHeaderAttachmentDescription, workSheetName, colIndexAttachmentDescription + 1, "Attachment Description");
				this.ValidateCellHeaderName(cellHeaderIsHaveVendorTaxId, workSheetName, colIndexIsHaveVendorTaxId + 1, "Is Have Vendor Tax Id");
				this.ValidateCellHeaderName(cellHeaderTaxIdNo, workSheetName, colIndexTaxIdNo + 1, "Tax Id No");
				this.ValidateCellHeaderName(cellHeaderTaxIdName, workSheetName, colIndexTaxIdName + 1, "Tax Id Name");
				this.ValidateCellHeaderName(cellHeaderTaxIdAddress, workSheetName, colIndexTaxIdAddress + 1, "Tax Id Address");
				this.ValidateCellHeaderName(cellHeaderTaxIdAttachmentPath, workSheetName, colIndexTaxIdAttachmentPath + 1, "Tax Id Attachment Path");
				this.ValidateCellHeaderName(cellHeaderTaxIdRemarks, workSheetName, colIndexTaxIdRemarks + 1, "Tax Id Remarks");
				#endregion
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexTaxIdRemarks))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						VendorDatabase param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							VendorName = this.TryGetString(currentRow.GetCell(colIndexVendorName), workSheetName, cellHeaderVendorName.StringCellValue, rowNum),
							VendorType = this.TryGetString(currentRow.GetCell(colIndexVendorType), workSheetName, cellHeaderVendorType.StringCellValue, rowNum),
							Country = this.TryGetString(currentRow.GetCell(colIndexCountry), workSheetName, cellHeaderCountry.StringCellValue, rowNum),
							City = this.TryGetString(currentRow.GetCell(colIndexCity), workSheetName, cellHeaderCity.StringCellValue, rowNum),
							Address = this.TryGetString(currentRow.GetCell(colIndexAddress), workSheetName, cellHeaderAddress.StringCellValue, rowNum),
							Phone1 = this.TryGetString(currentRow.GetCell(colIndexPhone1), workSheetName, cellHeaderPhone1.StringCellValue, rowNum),
							Phone2 = this.TryGetString(currentRow.GetCell(colIndexPhone2), workSheetName, cellHeaderPhone2.StringCellValue, rowNum),
							Fax = this.TryGetString(currentRow.GetCell(colIndexFax), workSheetName, cellHeaderFax.StringCellValue, rowNum),
							Email = this.TryGetString(currentRow.GetCell(colIndexEmail), workSheetName, cellHeaderEmail.StringCellValue, rowNum),
							IsHaveCertificateOfDomicile = this.TryGetBoolean(currentRow.GetCell(colIndexIsHaveCertificateofDomicile), workSheetName, cellHeaderIsHaveCertificateofDomicile.StringCellValue, rowNum),
							CoDAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexCoDAttachmentPath), workSheetName, cellHeaderCoDAttachmentPath.StringCellValue, rowNum),
							CoDExpiryDate = this.TryGetString(currentRow.GetCell(colIndexCoDExpiryDate), workSheetName, cellHeaderCoDExpiryDate.StringCellValue, rowNum),
							IsTaxableEntrepreneur = this.TryGetBoolean(currentRow.GetCell(colIndexIsTaxableEntrepreneur), workSheetName, cellHeaderIsTaxableEntrepreneur.StringCellValue, rowNum),
							SPPKPPath = this.TryGetString(currentRow.GetCell(colIndexSPPKPPath), workSheetName, cellHeaderSPPKPPath.StringCellValue, rowNum),
							IsUnderWithholdingTax = this.TryGetBoolean(currentRow.GetCell(colIndexIsUnderWithholdingTax), workSheetName, cellHeaderIsUnderWithholdingTax.StringCellValue, rowNum),
							Remarks = this.TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum),
							AttachmentPath = this.TryGetString(currentRow.GetCell(colIndexAttacmentPath), workSheetName, cellHeaderAttacmentPath.StringCellValue, rowNum),
							AttachmentDescription = this.TryGetString(currentRow.GetCell(colIndexAttachmentDescription), workSheetName, cellHeaderAttachmentDescription.StringCellValue, rowNum),
							IsHaveVendorTaxId = this.TryGetBoolean(currentRow.GetCell(colIndexIsHaveVendorTaxId), workSheetName, cellHeaderIsHaveVendorTaxId.StringCellValue, rowNum),
							TaxIdNo = this.TryGetString(currentRow.GetCell(colIndexTaxIdNo), workSheetName, cellHeaderTaxIdNo.StringCellValue, rowNum),
							TaxIdName = this.TryGetString(currentRow.GetCell(colIndexTaxIdName), workSheetName, cellHeaderTaxIdName.StringCellValue, rowNum),
							TaxIdAddress = this.TryGetString(currentRow.GetCell(colIndexTaxIdAddress), workSheetName, cellHeaderTaxIdAddress.StringCellValue, rowNum),
							TaxIdAttachmentPath = this.TryGetString(currentRow.GetCell(colIndexTaxIdAttachmentPath), workSheetName, cellHeaderTaxIdAttachmentPath.StringCellValue, rowNum),
							TaxIdRemarks = this.TryGetString(currentRow.GetCell(colIndexTaxIdRemarks), workSheetName, cellHeaderTaxIdRemarks.StringCellValue, rowNum)
						};
												
						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}

				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.VendorDatabase} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<VendorBusinessNameParam> ReadVendorBusinessNameParam(IWorkbook wb)
		{
			List<VendorBusinessNameParam> result = [];

			try
			{
				string workSheetName = SheetConstant.VendorDatabaseBusinessName;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexBusinessNameCategory = colIndexSeq + 1;
				int colIndexBusinessName = colIndexBusinessNameCategory + 1;
				int colIndexBusinessNameType = colIndexBusinessName + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderBusinessNameCategory = firstRow.GetCell(colIndexBusinessNameCategory);
				ICell cellHeaderBusinessName = firstRow.GetCell(colIndexBusinessName);
				ICell cellHeaderBusinessNameType = firstRow.GetCell(colIndexBusinessNameType);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderBusinessNameCategory, workSheetName, colIndexBusinessNameCategory + 1, "Business Name Category");
				this.ValidateCellHeaderName(cellHeaderBusinessName, workSheetName, colIndexBusinessName + 1, "Business Name");
				this.ValidateCellHeaderName(cellHeaderBusinessNameType, workSheetName, colIndexBusinessNameType + 1, "Business Name Type");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;

					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexBusinessNameType))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						VendorBusinessNameParam param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							BusinessNameCategory = this.TryGetString(currentRow.GetCell(colIndexBusinessNameCategory), workSheetName, cellHeaderBusinessNameCategory.StringCellValue, rowNum),
							BusinessName = this.TryGetString(currentRow.GetCell(colIndexBusinessName), workSheetName, cellHeaderBusinessName.StringCellValue, rowNum),
							BusinessNameType = this.TryGetString(currentRow.GetCell(colIndexBusinessNameType), workSheetName, cellHeaderBusinessNameType.StringCellValue, rowNum)
						};
						result.Add(param);
					}
					#endregion

				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.VendorDatabaseBusinessName} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<VendorBankAccountParam> ReadVendorBankAccountParam(IWorkbook wb)
		{
			List<VendorBankAccountParam> result = [];

			try
			{
				string workSheetName = SheetConstant.VendorDatabaseBankAccount;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexBankAccountCountry = colIndexSeq + 1;
				int colIndexBankAccountAliasName = colIndexBankAccountCountry + 1;
				int colIndexBankAccountBankName = colIndexBankAccountAliasName + 1;
				int colIndexBankAccountBranch = colIndexBankAccountBankName + 1;
				int colIndexBankAccountNo = colIndexBankAccountBranch + 1;
				int colIndexBankAccountName = colIndexBankAccountNo + 1;
				int colIndexBankAccountCurrency = colIndexBankAccountName + 1;
				int colIndexBankAccountSwiftCode = colIndexBankAccountCurrency + 1;
				int colIndexBankAccountIsDefault = colIndexBankAccountSwiftCode + 1;
				int colIndexBankAccountEvidencePath = colIndexBankAccountIsDefault + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderBankAccountCountry = firstRow.GetCell(colIndexBankAccountCountry);
				ICell cellHeaderBankAccountAliasName = firstRow.GetCell(colIndexBankAccountAliasName);
				ICell cellHeaderBankAccountBankName = firstRow.GetCell(colIndexBankAccountBankName);
				ICell cellHeaderBankAccountBranch = firstRow.GetCell(colIndexBankAccountBranch);
				ICell cellHeaderBankAccountNo = firstRow.GetCell(colIndexBankAccountNo);
				ICell cellHeaderBankAccountName = firstRow.GetCell(colIndexBankAccountName);
				ICell cellHeaderBankAccountCurrency = firstRow.GetCell(colIndexBankAccountCurrency);
				ICell cellHeaderBankAccountSwiftCode = firstRow.GetCell(colIndexBankAccountSwiftCode);
				ICell cellHeaderBankAccountIsDefault = firstRow.GetCell(colIndexBankAccountIsDefault);
				ICell cellHeaderBankAccountEvidencePath = firstRow.GetCell(colIndexBankAccountEvidencePath);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderBankAccountCountry, workSheetName, colIndexBankAccountCountry + 1, "Bank Account Country");
				this.ValidateCellHeaderName(cellHeaderBankAccountAliasName, workSheetName, colIndexBankAccountAliasName + 1, "Bank Account Alias Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountBankName, workSheetName, colIndexBankAccountBankName + 1, "Bank Account Bank Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountBranch, workSheetName, colIndexBankAccountBranch + 1, "Bank Account Branch");
				this.ValidateCellHeaderName(cellHeaderBankAccountNo, workSheetName, colIndexBankAccountNo + 1, "Bank Account No");
				this.ValidateCellHeaderName(cellHeaderBankAccountName, workSheetName, colIndexBankAccountName + 1, "Bank Account Name");
				this.ValidateCellHeaderName(cellHeaderBankAccountCurrency, workSheetName, colIndexBankAccountCurrency + 1, "Bank Account Currency");
				this.ValidateCellHeaderName(cellHeaderBankAccountSwiftCode, workSheetName, colIndexBankAccountSwiftCode + 1, "Bank Account Swift Code");
				this.ValidateCellHeaderName(cellHeaderBankAccountIsDefault, workSheetName, colIndexBankAccountIsDefault + 1, "Bank Account Is Default");
				this.ValidateCellHeaderName(cellHeaderBankAccountEvidencePath, workSheetName, colIndexBankAccountEvidencePath + 1, "Bank Account Evidence Path");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;

					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexBankAccountEvidencePath))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						VendorBankAccountParam param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							BankAccountCountry = this.TryGetString(currentRow.GetCell(colIndexBankAccountCountry), workSheetName, cellHeaderBankAccountCountry.StringCellValue, rowNum),
							BankAccountAliasName = this.TryGetString(currentRow.GetCell(colIndexBankAccountAliasName), workSheetName, cellHeaderBankAccountAliasName.StringCellValue, rowNum),
							BankAccountBankName = this.TryGetString(currentRow.GetCell(colIndexBankAccountBankName), workSheetName, cellHeaderBankAccountBankName.StringCellValue, rowNum),
							BankAccountBranch = this.TryGetString(currentRow.GetCell(colIndexBankAccountBranch), workSheetName, cellHeaderBankAccountBranch.StringCellValue, rowNum),
							BankAccountNo = this.TryGetString(currentRow.GetCell(colIndexBankAccountNo), workSheetName, cellHeaderBankAccountNo.StringCellValue, rowNum),
							BankAccountName = this.TryGetString(currentRow.GetCell(colIndexBankAccountName), workSheetName, cellHeaderBankAccountName.StringCellValue, rowNum),
							BankAccountCurrency = this.TryGetString(currentRow.GetCell(colIndexBankAccountCurrency), workSheetName, cellHeaderBankAccountCurrency.StringCellValue, rowNum),
							BankAccountSwiftCode = this.TryGetString(currentRow.GetCell(colIndexBankAccountSwiftCode), workSheetName, cellHeaderBankAccountSwiftCode.StringCellValue, rowNum),
							BankAccountIsDefault = this.TryGetBoolean(currentRow.GetCell(colIndexBankAccountIsDefault), workSheetName, cellHeaderBankAccountIsDefault.StringCellValue, rowNum),
							BankAccountEvidencePath = this.TryGetString(currentRow.GetCell(colIndexBankAccountEvidencePath), workSheetName, cellHeaderBankAccountEvidencePath.StringCellValue, rowNum)
						};
						result.Add(param);
					}
					#endregion

				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.VendorDatabaseBankAccount} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<VendorContactPersonParam> ReadVendorContactPersonParam(IWorkbook wb)
		{
			List<VendorContactPersonParam> result = [];

			try
			{
				string workSheetName = SheetConstant.VendorDatabaseContactPerson;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexContactPersonName = colIndexSeq + 1;
				int colIndexContactPersonPhone1 = colIndexContactPersonName + 1;
				int colIndexContactPersonPhone2 = colIndexContactPersonPhone1 + 1;
				int colIndexContactPersonEmail = colIndexContactPersonPhone2 + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderContactPersonName = firstRow.GetCell(colIndexContactPersonName);
				ICell cellHeaderContactPersonPhone1 = firstRow.GetCell(colIndexContactPersonPhone1);
				ICell cellHeaderContactPersonPhone2 = firstRow.GetCell(colIndexContactPersonPhone2);
				ICell cellHeaderContactPersonEmail = firstRow.GetCell(colIndexContactPersonEmail);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderContactPersonName, workSheetName, colIndexContactPersonName + 1, "Contact Person Name");
				this.ValidateCellHeaderName(cellHeaderContactPersonPhone1, workSheetName, colIndexContactPersonPhone1 + 1, "Contact Person Phone 1");
				this.ValidateCellHeaderName(cellHeaderContactPersonPhone2, workSheetName, colIndexContactPersonPhone2 + 1, "Contact Person Phone 2");
				this.ValidateCellHeaderName(cellHeaderContactPersonEmail, workSheetName, colIndexContactPersonEmail + 1, "Contact Person Email");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;

					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexContactPersonEmail))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string strSeq = this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						VendorContactPersonParam param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							DataFor = this.TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum),
							Sequence = !string.IsNullOrEmpty(strSeq) ? int.Parse(strSeq) : 1,
							ContactPersonName = this.TryGetString(currentRow.GetCell(colIndexContactPersonName), workSheetName, cellHeaderContactPersonName.StringCellValue, rowNum),
							ContactPersonPhone1 = this.TryGetString(currentRow.GetCell(colIndexContactPersonPhone1), workSheetName, cellHeaderContactPersonPhone1.StringCellValue, rowNum),
							ContactPersonPhone2 = this.TryGetString(currentRow.GetCell(colIndexContactPersonPhone2), workSheetName, cellHeaderContactPersonPhone2.StringCellValue, rowNum),
							ContactPersonEmail = this.TryGetString(currentRow.GetCell(colIndexContactPersonEmail), workSheetName, cellHeaderContactPersonEmail.StringCellValue, rowNum)
						};
						result.Add(param);
					}
					#endregion

				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.VendorDatabaseContactPerson} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		private List<VendorDatabaseTask> ReadVendorDatabaseTask(IWorkbook wb, List<TestPlan> plans)
		{
			List<VendorDatabaseTask> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.VendorDatabase).ToList();
			try
			{
				string workSheetName = SheetConstant.VendorDatabaseTask;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexSeq = colIndexId + 1;
				int colIndexActor = colIndexSeq + 1;
				int colIndexAction = colIndexActor + 1;				
				int colIndexReturnTo = colIndexAction + 1;
				int colIndexNotes = colIndexReturnTo + 1;
				int colIndexNewVendorCode = colIndexNotes + 1;
				int colIndexResult = colIndexNewVendorCode + 1;
				int colIndexScreenshot = colIndexResult + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderActor = firstRow.GetCell(colIndexActor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderAction = firstRow.GetCell(colIndexAction);
				ICell cellHeaderReturnTo = firstRow.GetCell(colIndexReturnTo);
				ICell cellHeaderNotes = firstRow.GetCell(colIndexNotes);
				ICell cellHeaderNewVendorCode = firstRow.GetCell(colIndexNewVendorCode);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				this.ValidateCellHeaderName(cellHeaderReturnTo, workSheetName, colIndexReturnTo + 1, "Return To");
				this.ValidateCellHeaderName(cellHeaderNotes, workSheetName, colIndexNotes + 1, "Notes");
				this.ValidateCellHeaderName(cellHeaderNewVendorCode, workSheetName, colIndexNewVendorCode + 1, "New Vendor Code");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexScreenshot))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						VendorDatabaseTask param = new()
						{
							Row = rowNum,
							TestCaseId = int.Parse(id),
							Sequence = int.Parse(this.TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum)),
							Actor = this.TryGetString(currentRow.GetCell(colIndexActor), workSheetName, cellHeaderActor.StringCellValue, rowNum),
							Action = this.TryGetString(currentRow.GetCell(colIndexAction), workSheetName, cellHeaderAction.StringCellValue, rowNum)
						};
						param.ReturnTo = this.TryGetString(currentRow.GetCell(colIndexReturnTo), workSheetName, cellHeaderReturnTo.StringCellValue, rowNum);
						param.Notes = this.TryGetString(currentRow.GetCell(colIndexNotes), workSheetName, cellHeaderNotes.StringCellValue, rowNum);
						param.NewVendorCode = this.TryGetString(currentRow.GetCell(colIndexNewVendorCode), workSheetName, cellHeaderNewVendorCode.StringCellValue, rowNum);


						if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
							result.Add(param);
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.VendorDatabaseTask} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}
		
		public List<TestPlan> ReadTestPlans(IWorkbook wb)
		{
			List<TestPlan> result = [];
			try
			{
				string workSheetName = "TestPlan";
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexTestCaseId = 0;
				int colIndexModuleName = colIndexTestCaseId + 1;
				int colIndexTestCase = colIndexModuleName + 1;
				int colIndexFromTestCaseId = colIndexTestCase + 1;
				int colIndexStatus = colIndexFromTestCaseId + 1;
				int colIndexDateTest = colIndexStatus + 1;
				int colIndexRemarks = colIndexDateTest + 1;
				int colIndexScreenCapture = colIndexRemarks + 1;
				int colIndexTestData = colIndexScreenCapture + 1;
				int colIndexReqNum = colIndexTestData + 1;
				int colIndexUserLogin = colIndexReqNum + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderTestCaseId = firstRow.GetCell(colIndexTestCaseId);
				ICell cellHeaderModuleName = firstRow.GetCell(colIndexModuleName);
				ICell cellHeaderTestCase = firstRow.GetCell(colIndexTestCase);
				ICell cellHeaderFromTestCaseId = firstRow.GetCell(colIndexFromTestCaseId);
				ICell cellHeaderStatus = firstRow.GetCell(colIndexStatus);
				ICell cellHeaderDateTest = firstRow.GetCell(colIndexDateTest);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderScreenCapture = firstRow.GetCell(colIndexScreenCapture);
				ICell cellHeaderTestData = firstRow.GetCell(colIndexTestData);
				ICell cellHeaderReqNum = firstRow.GetCell(colIndexReqNum);
				ICell cellHeaderUserLogin = firstRow.GetCell(colIndexUserLogin);

				ValidateCellHeaderName(cellHeaderTestCaseId, workSheetName, colIndexTestCaseId + 1, "Test Case Id");
				ValidateCellHeaderName(cellHeaderModuleName, workSheetName, colIndexModuleName + 1, "Module Name");
				ValidateCellHeaderName(cellHeaderTestCase, workSheetName, colIndexTestCase + 1, "Test Case");
				ValidateCellHeaderName(cellHeaderStatus, workSheetName, colIndexStatus + 1, "Status");
				ValidateCellHeaderName(cellHeaderDateTest, workSheetName, colIndexDateTest + 1, "Date Test");
				ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				ValidateCellHeaderName(cellHeaderScreenCapture, workSheetName, colIndexScreenCapture + 1, "Screen Capture");
				ValidateCellHeaderName(cellHeaderTestData, workSheetName, colIndexTestData + 1, "Test Data");
				ValidateCellHeaderName(cellHeaderUserLogin, workSheetName, colIndexUserLogin + 1, "User Login");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexUserLogin))
					{
						break;
					}

					#region Assign Param
					string id = this.TryGetString(currentRow.GetCell(colIndexTestCaseId), workSheetName, cellHeaderTestCaseId.StringCellValue, rowNum);
					var status = this.TryGetString(currentRow.GetCell(colIndexStatus), workSheetName, cellHeaderStatus.StringCellValue, rowNum);
					string moduleName = this.TryGetString(currentRow.GetCell(colIndexModuleName), workSheetName, cellHeaderModuleName.StringCellValue, rowNum);
					string testCase = this.TryGetString(currentRow.GetCell(colIndexTestCase), workSheetName, cellHeaderTestCase.StringCellValue, rowNum);
					string fromTesctCaseId = TryGetString(currentRow.GetCell(colIndexFromTestCaseId), workSheetName, cellHeaderFromTestCaseId.StringCellValue, rowNum);

					if (!string.IsNullOrEmpty(id))
					{
						if (!string.IsNullOrEmpty(moduleName) || !string.IsNullOrEmpty(testCase))
						{
							if (!moduleName.Equals(ModuleNameConstant.RejectBankConfirmation) && !string.IsNullOrEmpty(testCase))
							{
								TestPlan testPlan = new()
								{
									Row = rowNum,
									TestCaseId = int.Parse(id),
									FromTestCaseId = !string.IsNullOrEmpty(fromTesctCaseId) ? int.Parse(fromTesctCaseId) : 0,
									ModuleName = moduleName,
									TestCase = testCase,
									Status = status,
									RequestResult = TryGetString(currentRow.GetCell(colIndexReqNum), workSheetName, cellHeaderReqNum.StringCellValue, rowNum),
									UserLogin = this.TryGetString(currentRow.GetCell(colIndexUserLogin), workSheetName, cellHeaderUserLogin.StringCellValue, rowNum)
								};
								result.Add(testPlan);
							}
							if (moduleName.Equals(ModuleNameConstant.RejectBankConfirmation))
							{
								TestPlan testPlan = new()
								{
									Row = rowNum,
									TestCaseId = int.Parse(id),
									ModuleName = moduleName,
									TestCase = testCase,
									Status = status,
									UserLogin = this.TryGetString(currentRow.GetCell(colIndexUserLogin), workSheetName, cellHeaderUserLogin.StringCellValue, rowNum)
								};
								result.Add(testPlan);
							}
						}
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read {SheetConstant.TestPlan} is Fail : {AutomationHelper.CheckErrorReadExcel(ex.Message)}");
			}
		}

		public List<User> ReadUsers(IWorkbook wb)
		{
			List<User> result = [];
			try
			{
				string workSheetName = "User";
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexUsername = colIndexId + 1;
				int colIndexPassword = colIndexUsername + 1;
				int colIndexRole = colIndexPassword + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderUsername = firstRow.GetCell(colIndexUsername);
				ICell cellHeaderPassword = firstRow.GetCell(colIndexPassword);
				ICell cellHeaderRole = firstRow.GetCell(colIndexRole);

				this.ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Id");
				this.ValidateCellHeaderName(cellHeaderUsername, workSheetName, colIndexUsername + 1, "Username");
				this.ValidateCellHeaderName(cellHeaderPassword, workSheetName, colIndexPassword + 1, "Password");
				this.ValidateCellHeaderName(cellHeaderRole, workSheetName, colIndexRole + 1, "Role");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (this.IsRowEmpty(currentRow, colIndexRole))
					{
						break;
					}

					User user = new();
					#region Assign Param
					user.Id = int.Parse(this.TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum));
					user.Username = this.TryGetString(currentRow.GetCell(colIndexUsername), workSheetName, cellHeaderUsername.StringCellValue, rowNum);
					user.Password = this.TryGetString(currentRow.GetCell(colIndexPassword), workSheetName, cellHeaderPassword.StringCellValue, rowNum);
					user.Role = this.TryGetString(currentRow.GetCell(colIndexRole), workSheetName, cellHeaderRole.StringCellValue, rowNum);
					#endregion
					result.Add(user);
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Fail ReadUsers : {ex.Message}");
			}
		}

		#endregion

		#region Write
		public async void WriteApproval<T>(AppConfig cfg, T param, string moduleName) where T : class
		{
			try
			{
				string dateTimeNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

				using (var fs = new FileStream(cfg.ExcelConfigPath, FileMode.Open, FileAccess.Read))
				{
					var wb = new XSSFWorkbook(fs);
					ISheet sheet;
					int resultColumnIndex, captureColumnIndex, rowIndex;
					string testCase, status;

					if (param is VendorDatabaseTask vendorParam)
					{
						sheet = wb.GetSheet(SheetConstant.VendorDatabaseTask);
						resultColumnIndex = 7;
						captureColumnIndex = 8;
						VendorDatabaseTask newParam = param as VendorDatabaseTask ?? new();
						rowIndex = newParam.Row;
						testCase = $"{TestCaseConstant.Approval} - {newParam.Action}";
						status = newParam.Result == StatusConstant.Success ? StatusConstant.Success : newParam.Result;
					}
					else if (param is PVTask pvParam)
					{
						sheet = wb.GetSheet(SheetConstant.PVTask);
						resultColumnIndex = 7;
						captureColumnIndex = 8;
						PVTask newParam = param as PVTask ?? new();
						rowIndex = newParam.Row;
						testCase = $"{TestCaseConstant.Approval} - {newParam.Action}";
						status = newParam.Result == StatusConstant.Success ? StatusConstant.Success : newParam.Result;
					}
					else
					{
						throw new InvalidOperationException("Unsupported parameter type.");
					}

					var row = sheet.GetRow(rowIndex);
					ICell cellResult = row.CreateCell(resultColumnIndex);
					cellResult.CellStyle.Alignment = HorizontalAlignment.Left;
					cellResult.SetCellValue(status);

					TestPlan plan = new()
					{
						TestCase = testCase,
						ModuleName = moduleName,
						Status = status.Contains("Error") ? "Fail" : "Success"
					};

					string filePath = string.Empty;
					try
					{
						filePath = await ScreenCapture(cfg, plan);
					}
					catch (Exception ex)
					{
						Console.WriteLine($"Screen Capture failed: {ex.Message}");
					}

					if (!string.IsNullOrEmpty(filePath))
					{
						var cellScreenCapture = row.CreateCell(captureColumnIndex);
						var creationHelper = wb.GetCreationHelper();
						var hyperlink = creationHelper.CreateHyperlink(HyperlinkType.File);
						hyperlink.Address = filePath;
						cellScreenCapture.Hyperlink = hyperlink;
						cellScreenCapture.CellStyle.Alignment = HorizontalAlignment.Center;
						cellScreenCapture.SetCellValue("Open File");
					}

					using (var outputStream = new FileStream(cfg.ExcelConfigPath, FileMode.Create, FileAccess.Write))
					{
						wb.Write(outputStream);
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Fail WriteApproval: {ex.Message}");
			}
		}

		public async void WriteAutomationResult(AppConfig cfg, TestPlan param)
		{
			try
			{
				string dateNow = DateTime.Now.ToString("yyyy/MM/dd");
				string dateTimeNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

				using (var fs = new FileStream(cfg.ExcelConfigPath, FileMode.Open, FileAccess.Read))
				{
					var wb = new XSSFWorkbook(fs);
					ISheet sheet = wb.GetSheet("TestPlan");

					IRow row = sheet.GetRow(param.Row);
					var cell4 = row.CreateCell(4);
					cell4.SetCellValue(param.Status);
					cell4.CellStyle.Alignment = HorizontalAlignment.Center;
					var cell5 = row.CreateCell(5);
					cell5.SetCellValue(dateTimeNow);
					cell5.CellStyle.Alignment = HorizontalAlignment.Center;
					var cell6 = row.CreateCell(6);
					cell6.SetCellValue(param.Remarks);
					cell6.CellStyle.Alignment = HorizontalAlignment.Left;

					string filePath = string.Empty;
					try
					{
						filePath = await ScreenCapture(cfg, param);
					}
					catch (Exception ex)
					{
						Console.WriteLine($"Screen Capture failed: {ex.Message}");
					}

					if (!string.IsNullOrEmpty(filePath))
					{
						var cellScreenCapture = row.CreateCell(7);
						ICreationHelper creationHelper = wb.GetCreationHelper();
						IHyperlink hyperlink = creationHelper.CreateHyperlink(HyperlinkType.File);
						hyperlink.Address = filePath;
						cellScreenCapture.Hyperlink = hyperlink;
						cellScreenCapture.CellStyle.Alignment = HorizontalAlignment.Center;

						cellScreenCapture.SetCellValue("Open File");
					}

					if (!string.IsNullOrEmpty(param.TestData))
					{
						var testData = row.CreateCell(8);
						ICreationHelper creationHelper = wb.GetCreationHelper();
						IHyperlink hyperlink = creationHelper.CreateHyperlink(HyperlinkType.Document);
						hyperlink.Address = param.TestData;
						testData.Hyperlink = hyperlink;
						testData.CellStyle.Alignment = HorizontalAlignment.Center;

						testData.SetCellValue("Go To Data");
					}

					if (!string.IsNullOrEmpty(param.RequestResult))
					{
						var cell9 = row.CreateCell(9);
						cell9.SetCellValue(param.RequestResult);
						cell9.CellStyle.Alignment = HorizontalAlignment.Left;
					}

					using (var outputStream = new FileStream(cfg.ExcelConfigPath, FileMode.Create, FileAccess.Write))
					{
						wb.Write(outputStream);
						outputStream.Close();
					}
					fs.Close();
					wb.Close();
					Thread.Sleep(1500);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Fail Write Automation Result : {ex.Message}");
			}
		}
		#endregion

		#region Private Function

		bool AreAllPropertiesSet(PVDetailDetailTransactionModels detail)
		{
			return !string.IsNullOrWhiteSpace(detail.Date) &&
				   !string.IsNullOrWhiteSpace(detail.Description) &&
				   !string.IsNullOrWhiteSpace(detail.Amount) &&
				   !string.IsNullOrWhiteSpace(detail.Currency) &&
				   !string.IsNullOrWhiteSpace(detail.Rate) &&
				   !string.IsNullOrWhiteSpace(detail.AttachmentPath);
		}
		[DllImport("user32.dll")]
		private static extern int GetSystemMetrics(int nIndex);

		private const int SM_CXSCREEN = 0;
		private const int SM_CYSCREEN = 1;
		private async Task<string> ScreenCapture(AppConfig cfg, TestPlan param)
		{
			try
			{
				string dateNow = DateTime.Now.ToString("yyyy-MM-dd");
				string dateTime = DateTime.Now.ToString("HHmmss");
				var filePath = Path.Combine(dateNow, $"{dateTime}_{param.TestCase}_{param.ModuleName}_{param.Status}.png");
				string fullFilePath = Path.Combine(cfg.ScreenCapturePath, filePath);
				DirectoryCheck(fullFilePath);

				// Capture the screen
				int screenWidth = GetSystemMetrics(SM_CXSCREEN);
				int screenHeight = GetSystemMetrics(SM_CYSCREEN);

				using var bitmap = new Bitmap(screenWidth, screenHeight);
				using var graphics = Graphics.FromImage(bitmap);
				graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
				bitmap.Save(fullFilePath, ImageFormat.Png);

				return $"ScreenCapture\\{filePath}";
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Failed to get screen capture {param.ModuleName} - {param.TestCase}: {ex.Message}");
				return "";
			}
		}

		private bool ValidateCellHeaderName(ICell cellHeader, string worksheetName, int cellHeaderDisplayIndex, string validHeaderNames)
		{
			string cellHeaderValue = cellHeader.StringCellValue.Trim();

			if (cellHeaderValue.Equals(validHeaderNames, StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
			else
			{
				Console.WriteLine($"{worksheetName} - Column {cellHeaderDisplayIndex.ToString()} must be {validHeaderNames}");
				return false;
			}
		}

		private string TryGetString(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			string result = string.Empty;
			headerName = headerName.Trim();
			if (cell == null)
			{
				//Console.WriteLine($"{worksheetName} - Column {headerName} - Row {rowNum} is empty.");
			}
			else
			{
				return this.TryGetStringAllowNull(cell, worksheetName, headerName, rowNum) ?? string.Empty;
			}
			return result;
		}

		private string? TryGetStringAllowNull(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			string result = string.Empty;
			headerName = headerName.Trim();
			if (cell == null)
			{
				return null;
			}
			else
			{
				switch (cell.CellType)
				{
					case CellType.Numeric:
						result = cell.NumericCellValue.ToString();
						break;
					case CellType.String:
						result = cell.StringCellValue;
						break;
					case CellType.Boolean:
						result = cell.BooleanCellValue.ToString();
						break;
					case CellType.Unknown:
					case CellType.Formula:
					case CellType.Blank:
						result = cell.ToString() ?? string.Empty;
						break;
					case CellType.Error:
					default:
						Console.WriteLine($"{worksheetName} - An error occured when trying to get the value of Column {headerName} - Row {rowNum} is empty.");
						break;
				}
			}
			return result;
		}

		private bool IsRowEmpty(IRow row, int maxColumnIndex)
		{
			bool allColumnsEmpty = true;
			for (int colIndex = 0; colIndex <= maxColumnIndex; colIndex++)
			{
				if (row.GetCell(colIndex) != null && row.GetCell(colIndex).CellType != CellType.Blank)
				{
					allColumnsEmpty = false;
					break;
				}
			}

			return allColumnsEmpty;
		}

		private bool TryGetBoolean(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			bool result = false;
			headerName = headerName.Trim();
			if (cell == null)
			{
				//Console.WriteLine($"{worksheetName} - Column {headerName} - Row {rowNum} is empty.");
			}
			else
			{
				return this.TryGetBooleanAllowNull(cell, worksheetName, headerName, rowNum) ?? false;
			}
			return result;
		}
		private bool? TryGetBooleanAllowNull(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			bool? result = null;
			headerName = headerName.Trim();
			if (cell != null)
			{
				try
				{
					switch (cell.CellType)
					{
						case CellType.Numeric:
							result = Convert.ToBoolean(cell.NumericCellValue);
							break;
						case CellType.String:
							result = Convert.ToBoolean(cell.StringCellValue);
							break;
						case CellType.Boolean:
							result = cell.BooleanCellValue;
							break;
						case CellType.Unknown:
						case CellType.Formula:
						case CellType.Blank:
							break;
						case CellType.Error:
						default:
							Console.WriteLine($"{worksheetName} - An error occured when trying to get the value of Column {headerName} - Row {rowNum} is empty.");
							break;
					}
				}
				catch
				{
					Console.WriteLine($"{worksheetName} - An error occured when trying to get the value of Column {headerName} - Row {rowNum} is empty.");
				}
			}
			return result;
		}

		#endregion
	}
}
