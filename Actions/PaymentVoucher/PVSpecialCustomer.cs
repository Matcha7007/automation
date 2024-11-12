using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;

using OpenQA.Selenium.Edge;
using EPaymentVoucher.Locators.PaymentVoucher;
using OpenQA.Selenium;
using EPaymentVoucher.Models.General;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public interface IPVSpecialCustomer
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVSpecialCustomerModels param, bool isNew = true);
        void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVSpecialCustomerModels param);
    }
    public class PVSpecialCustomer(IHelperService utils) : IPVSpecialCustomer
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVSpecialCustomerModels pv = param.PVSpecialCustomers.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVSpecialCustomer, pv.Row);
                        HandleAddResubmit(driver, cfg, plan, pv);
                        break;
                    default: break;
                }
            }
            finally
            {
                driver.Dispose();
            }
        }

        public void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVSpecialCustomerModels param)
        {
            try
            {
                AutomationHelper.ClickElement(driver, PVSpecialCustomerXPath.ListRowEditDetail);
                if (!param.PaymentPurpose.Equals(PaymentPurposeSpecialCustomerConst.KasNegara, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElementNonMandatory(driver, PVSpecialCustomerDetailXPath.BankName, param.BankName);
                }
                AutomationHelper.FillElementNonMandatory(driver, PVSpecialCustomerDetailXPath.BankAccountNo, param.BankAccountNo);
                AutomationHelper.FillElementNonMandatory(driver, PVSpecialCustomerDetailXPath.BankAccountName, param.BankAccountName);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVSpecialCustomerModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.SpecialCustomer, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                AutomationHelper.ScrollToElement(driver, PVSpecialCustomerXPath.BtnAddListRequestorName);
                PVHelper.SearchRequestorOrEmployee(driver, PVSpecialCustomerXPath.BtnAddListRequestorName, param.RequestorName);

                HandleDetail(driver, param.PaymentPurpose, param, isNew);

                //if (!param.PaymentPurpose.Equals(PaymentPurposeSpecialCustomerConst.Other))
                //{
                //	try
                //	{
                //		PVHelper.FillOtherDetailPage(driver, pVBase, isNew, true);
                //	}
                //	catch (Exception ex) { throw new Exception($"The payment type {param.PaymentType} with the purpose {param.PaymentPurpose} requires an attachment in the other details section. Error : {ex.Message}"); }
                //}
                //else
                //{
                //}
                PVHelper.FillOtherDetailPage(driver, pVBase, isNew);

                if (isNew)
                {
                    string msg = AutomationHelper.GetAlertMessage(driver);
                    if (AutomationHelper.ValidateAlert(msg, "has been submited"))
                    {
                        plan.Status = StatusConstant.Success;
                        plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
                    }
                    else
                    {
                        throw new Exception(msg);
                    }
                }
                isAlert = true;
            }
            catch (Exception ex)
            {
                plan.Status = StatusConstant.Fail;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : {ex.Message}";
            }
            finally
            {
                if (isNew)
                {
					if (isAlert)
					{
						AutomationHelper.HandleAlertJS(driver, true);
						plan.RequestResult = PVHelper.SearchRPVNo(driver, param.Title);
					}
					utils.WriteAutomationResult(cfg, plan);
                    Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
                }
            }
        }

        public static void HandleDetail(EdgeDriver driver, string purpose, PVSpecialCustomerModels param, bool isNew)
        {
            try
            {
                if (!isNew)
                {
                    var listData = driver.FindElements(By.XPath(PVSpecialCustomerXPath.ListBtnDelete));
                    foreach (var item in listData)
                    {
                        item.Click();
                        AutomationHelper.HandleAlertJS(driver, true);
                    }
                    Thread.Sleep(1000);
                }

                AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerXPath.BtnAddDetail);

                AutomationHelper.FillElement(driver, PVSpecialCustomerDetailXPath.Payee, param.Payee);
                if (!param.PaymentPurpose.Equals(PaymentPurposeSpecialCustomerConst.KasNegara, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElement(driver, PVSpecialCustomerDetailXPath.PayeeType, param.PayeeType);
                    AutomationHelper.SelectElement(driver, PVSpecialCustomerDetailXPath.Country, param.Country);

                    if (param.Country.Equals("Indonesia", StringComparison.OrdinalIgnoreCase))
                    {
                        if (param.IsTaxable)
                        {
                            AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailXPath.TaxableTrue);
                            AutomationHelper.UploadFile(driver, PVSpecialCustomerDetailXPath.BtnSPPKPAttachment, param.SKPPPath);
                        }
                        else
                        {
                            AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailXPath.TaxableFalse);
                        }
                    }
                    else
                    {
                        if (param.IsCoD)
                        {
                            AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailXPath.CoDTrue);
                            AutomationHelper.UploadFile(driver, PVSpecialCustomerDetailXPath.BtnCoDAttachment, param.CoDPath);
                            AutomationHelper.DatePicker(driver, PVSpecialCustomerDetailXPath.CoDExpDate, param.CoDExpiryDate);
                        }
                        else
                        {
                            AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailXPath.CoDFalse);
                        }
                    }

                    if (param.IsUWTax)
                    {
                        AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailXPath.UWTaxTrue);
                    }
                    else
                    {
                        AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailXPath.UWTaxFalse);
                    }

                    if (!string.IsNullOrEmpty(param.TaxIdNo))
                    {
                        AutomationHelper.FillTaxId(driver, PVSpecialCustomerDetailXPath.TaxID, param.TaxIdNo);
                        AutomationHelper.FillElement(driver, PVSpecialCustomerDetailXPath.TaxIDName, param.TaxIdName);
                        AutomationHelper.FillElement(driver, PVSpecialCustomerDetailXPath.TaxIDAddress, param.TaxIdAddress);
                        AutomationHelper.UploadFile(driver, PVSpecialCustomerDetailXPath.BtnTaxIDAttachment, param.TaxIdAttachmentPath);
                    }
                    if (param.Country.Equals("Indonesia", StringComparison.OrdinalIgnoreCase))
                    {
                        AutomationHelper.SelectElement(driver, PVSpecialCustomerDetailXPath.BankNameSelect, param.BankName);
                    }
                    else
                    {
                        AutomationHelper.FillElement(driver, PVSpecialCustomerDetailXPath.BankName, param.BankName);
                    }
                }

                AutomationHelper.ScrollToElement(driver, PVSpecialCustomerDetailXPath.BtnSavePopUp);
                AutomationHelper.FillElement(driver, PVSpecialCustomerDetailXPath.BankAccountNo, param.BankAccountNo);
                AutomationHelper.FillElement(driver, PVSpecialCustomerDetailXPath.BankAccountName, param.BankAccountName);
                AutomationHelper.FillElementNonMandatory(driver, PVSpecialCustomerDetailXPath.Branch, param.Branch);

                try
                {
                    AutomationHelper.ClickElement(driver, PVSpecialCustomerDetailXPath.BtnAddDetail);
                }
                catch (Exception ex) { throw new Exception($"Unable to find the 'Add' button. Error : {ex.Message}"); }

                // detail
                PVHelper.SearchBusinessName(driver, PVSpecialCustomerDetailDetailXPath.BtnBusinessName, param.Detail.BussinessName);
                if (!param.PaymentPurpose.Equals(PaymentPurposeSpecialCustomerConst.KasNegara, StringComparison.OrdinalIgnoreCase))
                    AutomationHelper.SelectElement(driver, PVSpecialCustomerDetailDetailXPath.CostCenterName, param.Detail.CostCenter);
                AutomationHelper.FillElement(driver, PVSpecialCustomerDetailDetailXPath.LongDescription, param.Detail.LongDesc);
                AutomationHelper.FillElement(driver, PVSpecialCustomerDetailDetailXPath.ShortDescription, param.Detail.ShortDesc);
                AutomationHelper.FillElement(driver, PVSpecialCustomerDetailDetailXPath.GrossAmount, param.Detail.GrossAmount, true);

                if (!param.PaymentPurpose.Equals(PaymentPurposeSpecialCustomerConst.KasNegara, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElement(driver, PVSpecialCustomerDetailDetailXPath.Currency, param.Detail.Currency);

                    if (param.Detail.IsTaxExemption)
                    {
                        AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailDetailXPath.TaxExTrue);
                        AutomationHelper.UploadFile(driver, PVSpecialCustomerDetailDetailXPath.BtnSKBAttachment, param.Detail.SKBPath);
                    }
                    else
                    {
                        AutomationHelper.ClickInteractableElement(driver, PVSpecialCustomerDetailDetailXPath.TaxExFalse);
                    }
                }

                AutomationHelper.UploadFile(driver, PVSpecialCustomerDetailDetailXPath.BtnAttachment, param.Detail.AttachmentPath);
                AutomationHelper.FillElementNonMandatory(driver, PVSpecialCustomerDetailDetailXPath.Remarks, param.Detail.Remarks);
                AutomationHelper.ClickElement(driver, PVSpecialCustomerDetailDetailXPath.BtnSavePopUp);

                Thread.Sleep(500);

                AutomationHelper.ClickElement(driver, PVSpecialCustomerDetailXPath.BtnSavePopUp);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
