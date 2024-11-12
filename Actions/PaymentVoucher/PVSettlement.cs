using EPaymentVoucher.Locators.PaymentVoucher;
using EPaymentVoucher.Models;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.General;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Utilities;

using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public interface IPVSettlement
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVSettlementModels param, bool isNew = true);
    }

    public class PVSettlement(IHelperService utils) : IPVSettlement
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVSettlementModels data = param.PVSettlements.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVSettlement, data.Row);
                        HandleAddResubmit(driver, cfg, plan, data);
                        break;
                    default: break;
                }
            }
            finally
            {
                driver.Dispose();
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVSettlementModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.Settlement, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                AutomationHelper.ScrollToElement(driver, PVSettlementXPath.PaymentVoucherNo);
                PVHelper.SearchRequestorOrEmployee(driver, PVSettlementXPath.BtnRequestorName, param.RequestorName);

                AutomationHelper.SelectElement(driver, PVSettlementXPath.PaymentVoucherNo, param.PaymentVoucherNo);

                if (!isNew)
                {
                    var listDetails = driver.FindElements(By.XPath(PVReimbursementXPath.ListRowDeleteDetail));
                    foreach (var detail in listDetails)
                    {
                        detail.Click();
                        AutomationHelper.HandleAlertJS(driver, true);
                    }
                }

                foreach (PVDetailTransactionModels item in param.AddDetails)
                {
                    AddDetail(driver, item, param.PaymentPurpose);
                    break;
                }

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

        private static void AddDetail(EdgeDriver driver, PVDetailTransactionModels param, string purpose)
        {
            try
            {
                switch (purpose)
                {
                    case PaymentPurposeSettlementConst.Training:
                    case PaymentPurposeSettlementConst.Travel:
					case PaymentPurposeSettlementConst.SRE:
						AutomationHelper.ScrollToElement(driver, PVSettlementXPath.BtnAddDetail);
                        AutomationHelper.ClickElement(driver, PVSettlementXPath.BtnAddDetail);
                        break;
                    default:
                        AutomationHelper.ScrollToElement(driver, PVSettlementXPath.BtnAddDetail2);
                        AutomationHelper.ClickElement(driver, PVSettlementXPath.BtnAddDetail2);
                        break;
                }
                Thread.Sleep(500);

                PVHelper.AddDetailTransactionSettlement(driver, param, purpose);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
