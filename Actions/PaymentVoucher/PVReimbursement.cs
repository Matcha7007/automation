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
    public interface IPVReimbursement
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVReimbursementModels param, bool isNew = true);
        void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVReimbursementModels param);
    }

    public class PVReimbursement(IHelperService utils) : IPVReimbursement
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVReimbursementModels pvReimbursement = param.PVReimbursements.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVReimbursement, pvReimbursement.Row);
                        HandleAddResubmit(driver, cfg, plan, pvReimbursement);
                        break;
                    default: break;
                }
            }
            finally
            {
                driver.Dispose();
            }
        }

        public void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVReimbursementModels param)
        {
            try
            {
                AutomationHelper.SelectElementNonMandatory(driver, PVReimbursementXPath.BankName, param.BankName);
                AutomationHelper.FillElementNonMandatory(driver, PVReimbursementXPath.RequestorBankAccountName, param.RequestorBankAccountName);
                AutomationHelper.FillElementNonMandatory(driver, PVReimbursementXPath.RequestorBankAccountNo, param.RequestorBankAccountNo);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVReimbursementModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.Reimbursement, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                AutomationHelper.ScrollToElement(driver, PVReimbursementXPath.RequestorBankAccountNo);
                PVHelper.SearchRequestorOrEmployee(driver, PVReimbursementXPath.BtnAddListRequestorName, param.RequestorName);

                AutomationHelper.SelectElement(driver, PVReimbursementXPath.BankName, param.BankName);

                IWebElement requestorBankName = driver.FindElement(By.XPath(PVReimbursementXPath.RequestorBankAccountName));
                IWebElement requestorBankNo = driver.FindElement(By.XPath(PVReimbursementXPath.RequestorBankAccountNo));

                if (!isNew)
                {
                    requestorBankName.Clear();
                    requestorBankNo.Clear();
                }

                requestorBankName.SendKeys(param.RequestorBankAccountName);
                requestorBankNo.SendKeys(param.RequestorBankAccountNo);

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
                    case PaymentPurposeReimbursementConst.Reimbursement:
                    case PaymentPurposeReimbursementConst.HREvent:
                    case PaymentPurposeReimbursementConst.OfficeMatter:
                    case PaymentPurposeReimbursementConst.GiftAndEntertainment:
                        AutomationHelper.ScrollToElement(driver, PVReimbursementXPath.BtnAddDetail2);
                        AutomationHelper.ClickElement(driver, PVReimbursementXPath.BtnAddDetail2);
                        break;
                    default:
                        AutomationHelper.ScrollToElement(driver, PVReimbursementXPath.BtnAddDetail);
                        AutomationHelper.ClickElement(driver, PVReimbursementXPath.BtnAddDetail);
                        break;
                }
                Thread.Sleep(500);

                PVHelper.AddDetailTransaction(driver, param, purpose);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
