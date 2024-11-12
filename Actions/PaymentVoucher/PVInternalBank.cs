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
    public interface IPVInternalBank
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVInternalBankModels param, bool isNew = true);
        void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVInternalBankModels param);
    }
    public class PVInternalBank(IHelperService utils) : IPVInternalBank
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVInternalBankModels pv = param.PVInternalBanks.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVInternalBank, pv.Row);
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

        public void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVInternalBankModels param)
        {
            try
            {
                AutomationHelper.ClickElement(driver, PVInternalBankXPath.ListRowEditDetail);
                AutomationHelper.SelectElementNonMandatory(driver, PVInternalBankXPath.FromBankAccountNo, param.FromBankAccountNo);

                if (param.PaymentPurpose.Equals(PaymentPurposeInternalBankConst.InternalBank, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElementNonMandatory(driver, PVInternalBankXPath.ToBankAccountNo, param.ToBankAccountNo);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVInternalBankModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.InternalBank, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                HandleDetail(driver, param.PaymentPurpose, param, isNew);

                PVHelper.FillOtherDetailPage(driver, pVBase, isNew, true);

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

        public static void HandleDetail(EdgeDriver driver, string purpose, PVInternalBankModels param, bool isNew)
        {
            try
            {
                AutomationHelper.CheckElementExist(driver, PVInternalBankXPath.BtnAddDetail);
                if (!isNew)
                {
                    var listData = driver.FindElements(By.XPath(PVInternalBankXPath.ListBtnDelete));
                    foreach (var item in listData)
                    {
                        item.Click();
                        AutomationHelper.HandleAlertJS(driver, true);
                    }
                }

                AutomationHelper.ScrollToElement(driver, PVInternalBankXPath.BtnAddDetail);
                AutomationHelper.ClickElement(driver, PVInternalBankXPath.BtnAddDetail);

                AutomationHelper.SelectElement(driver, PVInternalBankXPath.FromBankAccountNo, param.FromBankAccountNo);

                if (purpose.Equals(PaymentPurposeInternalBankConst.InternalBank, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElement(driver, PVInternalBankXPath.ToBankAccountNo, param.ToBankAccountNo);
                }
                else
                {
                    AutomationHelper.SelectElement(driver, PVInternalBankXPath.CostCenter, param.CostCenter);
                }

                AutomationHelper.FillElement(driver, PVInternalBankXPath.LongDescription, param.LongDesc);
                AutomationHelper.FillElement(driver, PVInternalBankXPath.ShortDescription, param.ShortDesc);

                if (purpose.Equals(PaymentPurposeInternalBankConst.InternalBank, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElement(driver, PVInternalBankXPath.Currency, param.Currency);
                }

                AutomationHelper.FillElement(driver, PVInternalBankXPath.Amount, param.Amount, true);

                AutomationHelper.UploadFileNonMandatory(driver, PVInternalBankXPath.BtnAttachment, param.DetailAttachmentPath);
                AutomationHelper.FillElementNonMandatory(driver, PVInternalBankXPath.Remarks, param.Remarks);

                AutomationHelper.ClickElement(driver, PVInternalBankXPath.BtnSavePopUp);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
