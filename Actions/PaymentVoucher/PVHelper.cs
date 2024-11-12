using EPaymentVoucher.Locators.PaymentVoucher;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Utilities;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public static class PVHelper
    {
        public static void FillFirstPage(EdgeDriver driver, PVBaseModels param)
        {
            try
            {
                if (string.IsNullOrEmpty(param.PaymentPurpose))
                {
                    throw new Exception("Payment Purpose cannot be empty.");
                }
                AutomationHelper.CheckElementExist(driver, PVHeaderXPath.BtnNext);
                AutomationHelper.ScrollToElement(driver, PVHeaderXPath.BtnNext);
                AutomationHelper.FillField(driver, PVHeaderXPath.Title, param.Title);
                AutomationHelper.SelectElement(driver, PVHeaderXPath.PaymentType, param.PaymentType);
                AutomationHelper.SelectElement(driver, PVHeaderXPath.PaymentPurpose, param.PaymentPurpose);
                AutomationHelper.ClickElement(driver, PVHeaderXPath.BtnNext);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static string SearchRPVNo(EdgeDriver driver, string title)
        {
            try
            {
                Thread.Sleep(3000);
                AutomationHelper.FillElement(driver, PVListXPath.Title, title);
                AutomationHelper.ClickElement(driver, PVListXPath.BtnSearch);
                AutomationHelper.CheckElementExist(driver, PVListXPath.RowRequestNo);
                return driver.FindElement(By.XPath(PVListXPath.RowRequestNo)).Text;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fail Search RPV : {ex.Message}");
                return string.Empty;
            }
        }

        public static void FillOtherDetailPage(EdgeDriver driver, PVBaseModels param, bool isNew = true, bool isMandatoryAttachment = false)
        {
            try
            {
                if (isNew)
                {
                    AutomationHelper.ScrollToElement(driver, PVHeaderXPath.BtnNext);
                    AutomationHelper.ClickElement(driver, PVHeaderXPath.BtnNext);
                }

                AutomationHelper.ScrollToElement(driver, PVOtherDetailXPath.Remarks);
                if (!string.IsNullOrEmpty(param.AttachmentPath))
                {
                    if (!isNew)
                    {
                        var listAttachments = driver.FindElements(By.XPath(PVOtherDetailXPath.ListRowDeleteAttachment));
                        foreach (var attachment in listAttachments)
                        {
                            attachment.Click();
                            AutomationHelper.HandleAlertJS(driver, true);
                        }
                    }

                    if (isMandatoryAttachment)
                    {
                        if (string.IsNullOrEmpty(param.AttachmentPath) || string.IsNullOrEmpty(param.AttachmentDescription))
                        {
                            throw new Exception("Attachment and description on other detail pages are mandatory.");
                        }
                        AutomationHelper.UploadFile(driver, PVOtherDetailXPath.BtnAttachment, param.AttachmentPath);
                        AutomationHelper.FillElement(driver, PVOtherDetailXPath.Description, param.AttachmentDescription);
                        AutomationHelper.ClickElement(driver, PVOtherDetailXPath.BtnAddAttachment);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(param.AttachmentPath) && !string.IsNullOrEmpty(param.AttachmentDescription))
                        {
                            AutomationHelper.UploadFile(driver, PVOtherDetailXPath.BtnAttachment, param.AttachmentPath);
                            AutomationHelper.FillElement(driver, PVOtherDetailXPath.Description, param.AttachmentDescription);
                            AutomationHelper.ClickElement(driver, PVOtherDetailXPath.BtnAddAttachment);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(param.Remarks))
                {
                    IWebElement remarks = driver.FindElement(By.XPath(PVOtherDetailXPath.Remarks));
                    if (!isNew)
                        remarks.Clear();

                    remarks.SendKeys(param.Remarks);
                }
                Thread.Sleep(500);

                if (isNew)
                {
                    AutomationHelper.ScrollToElement(driver, PVOtherDetailXPath.BtnSubmit);
                    AutomationHelper.ClickElement(driver, PVOtherDetailXPath.BtnSubmit);
                }
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void SearchRequestorOrEmployee(EdgeDriver driver, string xpathAdd, string name)
        {
            try
            {
                AutomationHelper.ClickElement(driver, xpathAdd);
                AutomationHelper.FillElementNonMandatory(driver, PVPopUpEmployeeXPath.Requestor, name);
                AutomationHelper.ClickElement(driver, PVPopUpEmployeeXPath.BtnSearch);
                AutomationHelper.ClickElement(driver, PVPopUpEmployeeXPath.ListCellRequestor);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void SearchBusinessName(EdgeDriver driver, string xpathAdd, string name)
        {
            try
            {
                AutomationHelper.ClickElement(driver, xpathAdd);
                AutomationHelper.ClickElement(driver, PVPopUpBusinessNameXPath.BtnClear);
                AutomationHelper.ClickElement(driver, PVPopUpBusinessNameXPath.BtnSearch);
                AutomationHelper.FillElementNonMandatory(driver, PVPopUpBusinessNameXPath.Name, name);
                AutomationHelper.ClickElement(driver, PVPopUpBusinessNameXPath.BtnSearch);
                AutomationHelper.ClickInteractableElement(driver, PVPopUpBusinessNameXPath.ListCellCategory);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void SearchPayee(EdgeDriver driver, string xpathAdd, string name)
        {
            try
            {
                AutomationHelper.ClickElement(driver, xpathAdd);
                AutomationHelper.FillElementNonMandatory(driver, PVPopUpPayeeXPath.VendorName, name);
                AutomationHelper.ClickElement(driver, PVPopUpPayeeXPath.BtnSearch);
                AutomationHelper.ClickInteractableElement(driver, PVPopUpPayeeXPath.ListCellVendorCode);
                Thread.Sleep(500);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void AddDetailTransactionSettlement(EdgeDriver driver, PVDetailTransactionModels param, string purpose)
        {
            try
            {
                SearchBusinessName(driver, PVPopUpDetailTransactionXPath.BtnAddListBusinessName, param.BusinessName);

                AutomationHelper.SelectElement(driver, PVPopUpDetailTransactionXPath.CostCenter, param.CostCenter);

                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionXPath.LongDescription, param.LongDescription);
                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionXPath.ShortDescription, param.ShortDescription);

                if (purpose.Equals(PaymentPurposeConstant.Travel) || purpose.Equals(PaymentPurposeConstant.Training) || purpose.Equals(PaymentPurposeConstant.SRE))
                {
                    if (param.Details.Count > 0)
                    {
                        foreach (var item in param.Details)
                        {
                            AddDetailDetailTransaction(driver, item);
                        }
                    }
                    else
                    {
                        throw new Exception("There is no data available to provide details");
                    }
                }
                else
                {
                    AutomationHelper.DatePicker(driver, PVPopUpDetailTransactionXPath.Date, param.Date);
                    AutomationHelper.FillElement(driver, PVPopUpDetailTransactionXPath.Amount, param.Amount, true);
                    AutomationHelper.UploadFile(driver, PVPopUpDetailTransactionXPath.BtnAttachment, param.AttachmentPath);
                }

                AutomationHelper.ClickElement(driver, PVPopUpDetailTransactionXPath.BtnSave);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void EditSubmitDetailBankConfirm(EdgeDriver driver, PVDetailTransactionLongModels param)
        {
            try
            {
                AutomationHelper.SelectElementNonMandatory(driver, PVPopUpDetailTransactionLongXPath.BankAccountNo, param.BankAccountNo);
                AutomationHelper.ClickElement(driver, PVPopUpDetailTransactionLongXPath.BtnSavePopUp);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void AddDetailTransactionLong(EdgeDriver driver, PVDetailTransactionLongModels param, string purpose)
        {
            try
            {
                SearchRequestorOrEmployee(driver, PVPopUpDetailTransactionLongXPath.BtnRequestorName, param.RequestorName);

                AutomationHelper.SelectElement(driver, PVPopUpDetailTransactionLongXPath.CostCenterName, param.CostCenter);
                AutomationHelper.SelectElementNonMandatory(driver, PVPopUpDetailTransactionLongXPath.AccrualCode, param.AccrualCode);

                SearchBusinessName(driver, PVPopUpDetailTransactionLongXPath.BtnBusinessName, param.BusinessName);

                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionLongXPath.LongDescription, param.LongDescription);
                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionLongXPath.ShortDescription, param.ShortDescription);
                if (param.IsProposal)
                {
                    AutomationHelper.ClickInteractableElement(driver, PVPopUpDetailTransactionLongXPath.IsProposalTrue);
                    AutomationHelper.UploadFile(driver, PVPopUpDetailTransactionLongXPath.BtnProposal, param.ProposalPath);
                }
                else
                {
                    AutomationHelper.ClickInteractableElement(driver, PVPopUpDetailTransactionLongXPath.IsProposalFalse);
                }

                AutomationHelper.ScrollToElement(driver, PVPopUpDetailTransactionLongXPath.BtnSavePopUp);

                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionLongXPath.InvoiceNo, param.InvoiceNo);

                AutomationHelper.DatePicker(driver, PVPopUpDetailTransactionLongXPath.InvoiceDate, param.InvoiceDate);
                AutomationHelper.DatePicker(driver, PVPopUpDetailTransactionLongXPath.InvoiceDueDate, param.InvoiceDueDate);
                AutomationHelper.UploadFile(driver, PVPopUpDetailTransactionLongXPath.BtnInvoiceAttachment, param.InvoiceAttachmentPath);
                AutomationHelper.SelectElement(driver, PVPopUpDetailTransactionLongXPath.BankAccountNo, param.BankAccountNo);

                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionLongXPath.GrossAmount, param.GrossAmount, true);
                AutomationHelper.FillElementNonMandatory(driver, PVPopUpDetailTransactionLongXPath.Remarks, param.Remarks);

                AutomationHelper.ClickElement(driver, PVPopUpDetailTransactionLongXPath.BtnSavePopUp);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void AddDetailTransaction(EdgeDriver driver, PVDetailTransactionModels param, string purpose)
        {
            try
            {
                SearchBusinessName(driver, PVPopUpDetailTransactionXPath.BtnAddListBusinessName, param.BusinessName);

                AutomationHelper.SelectElement(driver, PVPopUpDetailTransactionXPath.CostCenter, param.CostCenter);
                AutomationHelper.SelectElementNonMandatory(driver, PVPopUpDetailTransactionXPath.AccrualCode, param.AccrualCode);
                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionXPath.LongDescription, param.LongDescription);
                AutomationHelper.FillElement(driver, PVPopUpDetailTransactionXPath.ShortDescription, param.ShortDescription);

                string xPathProposal = PVPopUpDetailTransactionXPath.BtnProposal;

                if (purpose.Equals(PaymentPurposeConstant.Travel) || purpose.Equals(PaymentPurposeConstant.Training))
                {
                    xPathProposal = xPathProposal.Replace("div[11]", "div[9]");
                    if (param.Details.Count > 0)
                    {
                        foreach (var item in param.Details)
                        {
                            AddDetailDetailTransaction(driver, item);
                            break;
                        }
                    }
                    else
                    {
                        throw new Exception("There is no data available to provide details");
                    }
                }
                else
                {
                    AutomationHelper.FillElement(driver, PVPopUpDetailTransactionXPath.Amount, param.Amount, true);
                    AutomationHelper.UploadFile(driver, PVPopUpDetailTransactionXPath.BtnAttachment, param.AttachmentPath);
                }
                if (param.IsProposal)
                {
                    AutomationHelper.ClickInteractableElement(driver, PVPopUpDetailTransactionXPath.IsProposalTrue);
                    AutomationHelper.UploadFile(driver, xPathProposal, param.ProposalPath);
                }
                else
                {
                    AutomationHelper.ClickInteractableElement(driver, PVPopUpDetailTransactionXPath.IsProposalFalse);
                }

                AutomationHelper.ClickElement(driver, PVPopUpDetailTransactionXPath.BtnSave);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void AddDetailDetailTransaction(EdgeDriver driver, PVDetailDetailTransactionModels param)
        {
            try
            {
                AutomationHelper.ClickElement(driver, PVPopUpDetailDetailTransactionXPath.BtnAddDetails);
                AutomationHelper.DatePicker(driver, PVPopUpDetailDetailTransactionXPath.Date, param.Date);
                AutomationHelper.SelectElement(driver, PVPopUpDetailDetailTransactionXPath.Description, param.Description);
                AutomationHelper.FillElement(driver, PVPopUpDetailDetailTransactionXPath.Amount, param.Amount, true);
                AutomationHelper.SelectElement(driver, PVPopUpDetailDetailTransactionXPath.Currency, param.Currency);
                AutomationHelper.FillElement(driver, PVPopUpDetailDetailTransactionXPath.Rate, param.Rate, true);
                AutomationHelper.UploadFile(driver, PVPopUpDetailDetailTransactionXPath.BtnAttachment, param.AttachmentPath);
                AutomationHelper.ClickElement(driver, PVPopUpDetailDetailTransactionXPath.BtnSave);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void AddRDDSalesForce(EdgeDriver driver, TestPlan plan, PVSalesForceModels param, bool isNew)
        {
            try
            {
                AutomationHelper.ScrollToElement(driver, PVSalesForceXPath.CostCenterName);
                AutomationHelper.SelectElement(driver, PVSalesForceXPath.CostCenterName, param.CostCenterName);

                if (param.PaymentPurpose.Equals(PaymentPurposeConstant.RDD) || param.PaymentPurpose.Equals(PaymentPurposeConstant.ContestAgency) || param.PaymentPurpose.Equals(PaymentPurposeConstant.ContestBankAssurance))
                    SearchBusinessName(driver, PVSalesForceXPath.BtnBusinessName, param.BusinessName);

                AutomationHelper.SelectElement(driver, PVSalesForceXPath.Payment, param.Payment);

                AutomationHelper.FillElement(driver, PVSalesForceXPath.ShortDescription, param.ShortDescription);

                bool isGross = param.Payment.Equals(PaymentConstant.Gross);

                HandleUploadExcel(driver, PVSalesForceXPath.BtnFile, PVSalesForceXPath.BtnUploadFile, param.FilePath, isGross);

                AutomationHelper.ScrollToElement(driver, PVSalesForceXPath.BtnFile);

                if (param.IsProposal)
                {
                    AutomationHelper.ClickInteractableElement(driver, PVSalesForceXPath.IsProposalTrue);
                    Thread.Sleep(500);
                    AutomationHelper.UploadFile(driver, PVSalesForceXPath.BtnProposal, param.ProposalPath);
                }
                else
                {
                    AutomationHelper.ClickInteractableElement(driver, PVSalesForceXPath.IsProposalFalse);
                }
                Thread.Sleep(500);

                if (plan.ModuleName.Equals(ModuleNameConstant.PVSalesForce) && !param.PaymentPurpose.Equals(PaymentPurposeConstant.CommissionAgency))
                    AutomationHelper.SelectElement(driver, PVSalesForceXPath.BeneficiaryCustomerType, param.BeneficiaryCustomerType);

                AutomationHelper.DatePicker(driver, PVSalesForceXPath.PeriodOfPayment, param.PeriodOfPayment);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }

        public static void HandleUploadExcel(EdgeDriver driver, string btnFile, string btnUpload, string path, bool isGross)
        {
            try
            {
                AutomationHelper.UploadFile(driver, btnFile, path);
                AutomationHelper.ClickElement(driver, btnUpload);

                string totalAmount = "";

                if (isGross)
                {
                    AutomationHelper.ScrollToElement(driver, PVSalesForceXPath.TotalGross);
                    Thread.Sleep(500);
                    totalAmount = driver.FindElement(By.XPath(PVSalesForceXPath.TotalGross)).Text;
                }
                else
                {
                    AutomationHelper.ScrollToElement(driver, PVSalesForceXPath.TotalNett);
                    Thread.Sleep(500);
                    totalAmount = driver.FindElement(By.XPath(PVSalesForceXPath.TotalNett)).Text;
                }

                if (totalAmount == "0.00")
                {
                    throw new Exception("Invalid Excel file");
                }
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
