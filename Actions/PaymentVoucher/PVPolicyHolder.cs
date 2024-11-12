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
    public interface IPVPolicyHolder
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVPolicyHolderModels param, bool isNew = true);
        void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVPolicyHolderModels param);
    }
    public class PVPolicyHolder(IHelperService utils) : IPVPolicyHolder
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVPolicyHolderModels pv = param.PVPolicyHolders.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVPolicyHolder, pv.Row);
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

        public void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVPolicyHolderModels param)
        {
            try
            {
                AutomationHelper.ClickElement(driver, PVPolicyHolderXPath.ListRowEditDetail);
                AutomationHelper.SelectElementNonMandatory(driver, PVPolicyHolderXPath.BankName, param.BankName);
                AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.BankAccountNo, param.BankAccountNo);
                AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.BankAccountName, param.BankAccountName);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVPolicyHolderModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.PolicyHolder, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                PVHelper.SearchRequestorOrEmployee(driver, PVPolicyHolderXPath.BtnAddListRequestorName, param.RequestorName);

                HandleDetail(driver, param.PaymentPurpose, param, isNew);

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

        public static void HandleDetail(EdgeDriver driver, string purpose, PVPolicyHolderModels param, bool isNew)
        {
            try
            {
                if (!isNew)
                {
                    var listData = driver.FindElements(By.XPath(PVPolicyHolderXPath.ListBtnDelete));
                    foreach (var item in listData)
                    {
                        item.Click();
                        AutomationHelper.HandleAlertJS(driver, true);
                    }
                    Thread.Sleep(500);
                }

                AutomationHelper.ClickElement(driver, PVPolicyHolderXPath.BtnAddDetail);    

                AutomationHelper.FillElement(driver, PVPolicyHolderXPath.Payee, param.Payee);
				if (purpose.Equals(PaymentPurposePolichHolderConst.NBUW, StringComparison.OrdinalIgnoreCase))
                {
					AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.PolicyNo, param.PolicyNo);
                }
                else
                {
					AutomationHelper.FillElement(driver, PVPolicyHolderXPath.PolicyNo, param.PolicyNo);
				}
				if (purpose.Equals(PaymentPurposePolichHolderConst.NBUW, StringComparison.OrdinalIgnoreCase)) AutomationHelper.FillElement(driver, PVPolicyHolderXPath.SPAJNo, param.SPAJNo);
				PVHelper.SearchBusinessName(driver, PVPolicyHolderXPath.BussinessName, param.BussinessName);
                AutomationHelper.FillElement(driver, PVPolicyHolderXPath.LongDesc, param.LongDesc);
                AutomationHelper.FillElement(driver, PVPolicyHolderXPath.ShortDesc, param.ShortDesc);
                AutomationHelper.SelectElement(driver, PVPolicyHolderXPath.Product, param.Product);
                AutomationHelper.SelectElement(driver, PVPolicyHolderXPath.Distribution, param.Distribution);
                AutomationHelper.SelectElement(driver, PVPolicyHolderXPath.Country, param.Country);
                AutomationHelper.SelectElement(driver, PVPolicyHolderXPath.BankName, param.BankName);
                AutomationHelper.FillElement(driver, PVPolicyHolderXPath.BankAccountNo, param.BankAccountNo);
                AutomationHelper.FillElement(driver, PVPolicyHolderXPath.BankAccountName, param.BankAccountName);
                AutomationHelper.ScrollToElement(driver, PVPolicyHolderXPath.BtnAttachment);
                AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.Branch, param.Branch);
                AutomationHelper.SelectElement(driver, PVPolicyHolderXPath.Currency, param.Currency);
                AutomationHelper.FillElement(driver, PVPolicyHolderXPath.GrossAmount, param.GrossAmount, true);

                if (!purpose.Equals(PaymentPurposePolichHolderConst.Claim, StringComparison.OrdinalIgnoreCase) && !purpose.Equals(PaymentPurposePolichHolderConst.NBUW, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.FillElement(driver, PVPolicyHolderXPath.ChargesAdmin, param.ChargesAdmin);
                }

                if (purpose.Equals(PaymentPurposePolichHolderConst.RefundPremiFreeLock, StringComparison.OrdinalIgnoreCase)|| purpose.Equals(PaymentPurposePolichHolderConst.NBUW, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.ChargesMedical, param.ChargesMedical);
                }

                AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.ForPremiumPayment, param.ForPremiumPayment);
                AutomationHelper.UploadFileNonMandatory(driver, PVPolicyHolderXPath.BtnAttachment, param.DetailAttachmentPath);
                AutomationHelper.ScrollToElement(driver, PVPolicyHolderXPath.PopupSave);

                if (!string.IsNullOrEmpty(param.TaxIdNo))
                {
                    AutomationHelper.FillTaxId(driver, PVPolicyHolderXPath.TaxIdNo, param.TaxIdNo);
                    AutomationHelper.FillElement(driver, PVPolicyHolderXPath.TaxIdName, param.TaxIdName);
                    AutomationHelper.FillElement(driver, PVPolicyHolderXPath.TaxIdAddress, param.TaxIdAddress);
                    AutomationHelper.UploadFile(driver, PVPolicyHolderXPath.BtnTaxIdAtaachmentPath, param.TaxIdAttachmentPath);

                }

                AutomationHelper.FillElementNonMandatory(driver, PVPolicyHolderXPath.Remarks, param.Remarks);
                AutomationHelper.ClickElement(driver, PVPolicyHolderXPath.PopupSave);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
