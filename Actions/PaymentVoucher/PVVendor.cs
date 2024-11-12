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
    public interface IPVVendor
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVVendorModels param, bool isNew = true);
        void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVVendorModels param);
    }

    public class PVVendor(IHelperService utils) : IPVVendor
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVVendorModels data = param.PVVendors.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVVendor, data.Row);
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

        public void HandleSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVVendorModels param)
        {
            try
            {
                AutomationHelper.ClickElement(driver, PVVendorXPath.ListRowEditDetail);

                foreach (PVDetailTransactionLongModels item in param.AddDetails)
                {
                    PVHelper.EditSubmitDetailBankConfirm(driver, item);
                    break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVVendorModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.Vendor, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                AutomationHelper.ScrollToElement(driver, PVVendorXPath.BtnPayee);
                PVHelper.SearchPayee(driver, PVVendorXPath.BtnPayee, param.Payee);

                if (param.IsRecurring)
                {
                    AutomationHelper.ClickInteractableElement(driver, PVVendorXPath.IsRecurringTrue);
                }
                else
                {
                    AutomationHelper.ClickInteractableElement(driver, PVVendorXPath.IsRecurringFalse);
                }

                if (param.PaymentPurpose.Equals(PaymentPurposeVendorConst.EProcurement, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.SelectElement(driver, PVVendorXPath.PO, param.PO);
                    AutomationHelper.SelectElement(driver, PVVendorXPath.GR, param.GR);
                    AutomationHelper.ClickElement(driver, PVVendorXPath.BtnAdd);

                    AutomationHelper.HandleAlertJS(driver, true);
                }
                else
                {
                    if (!isNew)
                    {
                        var listDetails = driver.FindElements(By.XPath(PVVendorXPath.ListRowDeleteDetail));
                        foreach (var detail in listDetails)
                        {
                            detail.Click();
                            AutomationHelper.HandleAlertJS(driver, true);
                        }
                    }

                    foreach (PVDetailTransactionLongModels item in param.AddDetails)
                    {
                        AddDetail(driver, item, param.PaymentPurpose);
                        break;
                    }
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

        private static void AddDetail(EdgeDriver driver, PVDetailTransactionLongModels param, string purpose)
        {
            try
            {
                AutomationHelper.ScrollToElement(driver, PVVendorXPath.BtnDetail);
                AutomationHelper.ClickElement(driver, PVVendorXPath.BtnDetail);
                PVHelper.AddDetailTransactionLong(driver, param, purpose);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
