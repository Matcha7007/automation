using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;

using OpenQA.Selenium.Edge;
using EPaymentVoucher.Locators.PaymentVoucher;
using EPaymentVoucher.Models.General;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public interface IPVBancassuranceBMI
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVBancassuranceBMIModels param, bool isNew = true);
    }
    public class PVBancassuranceBMI(IHelperService utils) : IPVBancassuranceBMI
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVBancassuranceBMIModels pv = param.PVBancassuranceBMIs.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVBancassuranceBMI, pv.Row);
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

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVBancassuranceBMIModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.BancassuranceBMI, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                HandleUpload(driver, param.PaymentPurpose, param, isNew);

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

        public static void HandleUpload(EdgeDriver driver, string purpose, PVBancassuranceBMIModels param, bool isNew)
        {
            try
            {
                AutomationHelper.ScrollToElement(driver, PVBancassuranceBMIXPath.CostCenter);
                if (purpose.Equals(PaymentPurposeBancassuranceBMIConst.WMABMI, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.UploadFile(driver, PVBancassuranceBMIXPath.BtnPaymentRelated, param.TaxReportWABMIPath);
                    AutomationHelper.UploadFileNonMandatory(driver, PVBancassuranceBMIXPath.BtnGLRelated, param.TemplateAdjustmentWABMIPath);
                }
                else if (purpose.Equals(PaymentPurposeBancassuranceBMIConst.FeebaseBMI, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.UploadFile(driver, PVBancassuranceBMIXPath.BtnPaymentRelated, param.TemplateFeebaseBMIPath);
                    AutomationHelper.UploadFile(driver, PVBancassuranceBMIXPath.BtnGLRelated, param.DetailBMIPath);
                }
                else if (purpose.Equals(PaymentPurposeBancassuranceBMIConst.CreditLifeBMI, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.UploadFile(driver, PVBancassuranceBMIXPath.BtnPaymentRelated, param.TemplateCreditLifeBMIPath);
                    AutomationHelper.UploadFileNonMandatory(driver, PVBancassuranceBMIXPath.BtnGLRelated, param.CreditLifeBMIPath);
                }
                else if (purpose.Equals(PaymentPurposeBancassuranceBMIConst.BankStaffBMI, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.UploadFile(driver, PVBancassuranceBMIXPath.BtnPaymentRelated, param.TaxCalculationBSPath);
                    AutomationHelper.UploadFileNonMandatory(driver, BankStafBMIXPath.BtnGLRelated, param.TemplateAdjustmentBankStaffBMIPath);
                    AutomationHelper.UploadFile(driver, BankStafBMIXPath.BtnDetailBMI, param.DetailBMIPath);
                }

                AutomationHelper.ClickElement(driver, PVBancassuranceBMIXPath.BtnUpload);
                string msg = AutomationHelper.GetAlertMessage(driver);
                if (AutomationHelper.ValidateAlert(msg, "success"))
                {
                    AutomationHelper.HandleAlertJS(driver, true);
                }
                else
                {
                    throw new Exception(msg);
                }
                AutomationHelper.ScrollToElement(driver, PVBancassuranceBMIXPath.CostCenter);

                AutomationHelper.SelectElement(driver, PVBancassuranceBMIXPath.CostCenter, param.CostCenterName);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
