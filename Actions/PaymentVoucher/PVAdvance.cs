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
    public interface IPVAdvance
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVAdvanceModels param, bool isNew = true);
        void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVAdvanceModels param);
    }
    public class PVAdvance(IHelperService utils) : IPVAdvance
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVAdvanceModels pvAdvance = param.PVAdvances.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVAdvance, pvAdvance.Row);
                        HandleAddResubmit(driver, cfg, plan, pvAdvance);
                        break;
                    default: break;
                }
            }
            finally
            {
                driver.Dispose();
            }
        }

        public void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVAdvanceModels param)
        {
            try
            {
                AutomationHelper.SelectElementNonMandatory(driver, PVAdvanceXPath.BankName, param.BankName);
                AutomationHelper.FillElementNonMandatory(driver, PVAdvanceXPath.RequestorBankAccountName, param.RequestorBankAccountName);
                AutomationHelper.FillElementNonMandatory(driver, PVAdvanceXPath.RequestorBankAccountNo, param.RequestorBankAccountNo);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVAdvanceModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.Advance, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                AutomationHelper.ScrollToElement(driver, PVAdvanceXPath.Reason);
                PVHelper.SearchRequestorOrEmployee(driver, PVAdvanceXPath.BtnRequestorName, param.RequestorName);

                AutomationHelper.SelectElement(driver, PVAdvanceXPath.BankName, param.BankName);

				AutomationHelper.FillElement(driver, PVAdvanceXPath.RequestorBankAccountName, param.RequestorBankAccountName);
				AutomationHelper.FillElement(driver, PVAdvanceXPath.RequestorBankAccountNo, param.RequestorBankAccountNo);

                AutomationHelper.DatePickerNonMandatory(driver, PVAdvanceXPath.StartDate, param.StartDate);
                AutomationHelper.DatePickerNonMandatory(driver, PVAdvanceXPath.EndDate, param.EndDate);

                AutomationHelper.ValidateNumeric(param.Amount);
				AutomationHelper.FillElement(driver, PVAdvanceXPath.Amount, param.Amount);
				AutomationHelper.FillElement(driver, PVAdvanceXPath.Reason, param.Reason);

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
    }
}
