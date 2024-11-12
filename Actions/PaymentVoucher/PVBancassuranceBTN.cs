using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;

using OpenQA.Selenium.Edge;
using EPaymentVoucher.Locators.PaymentVoucher;
using EPaymentVoucher.Models.General;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public interface IPVBancassuranceBTN
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVBancassuranceBTNModels param, bool isNew = true);
    }
    public class PVBancassuranceBTN(IHelperService utils) : IPVBancassuranceBTN
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVBancassuranceBTNModels pv = param.PVBancassuranceBTNs.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVBancassuranceBTN, pv.Row);
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

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVBancassuranceBTNModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.BancassuranceBTN, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
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

        public static void HandleUpload(EdgeDriver driver, string purpose, PVBancassuranceBTNModels param, bool isNew)
        {
            try
            {
                AutomationHelper.ScrollToElement(driver, PVBancassuranceBTNXPath.CostCenter);
                if (purpose.Equals(PaymentPurposeBancassuranceBTNConst.FeebaseBTN, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.UploadFile(driver, PVBancassuranceBTNXPath.BtnPaymentRelated, param.TemplateFeebaseBTNPath);
                    AutomationHelper.UploadFile(driver, PVBancassuranceBTNXPath.BtnGLRelated, param.DetailBTNPath);
                }
                else if (purpose.Equals(PaymentPurposeBancassuranceBTNConst.WMABTN, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.UploadFile(driver, PVBancassuranceBTNXPath.BtnPaymentRelated, param.TaxReportWMABTNPath);
                    AutomationHelper.UploadFileNonMandatory(driver, PVBancassuranceBTNXPath.BtnGLRelated, param.TemplateAdjustmentWMABTNPath);
                }

                AutomationHelper.ClickElement(driver, PVBancassuranceBTNXPath.BaseBtnUploadFile);
                string msg = AutomationHelper.GetAlertMessage(driver);
                if (AutomationHelper.ValidateAlert(msg, "success"))
                {
                    AutomationHelper.HandleAlertJS(driver, true);
                }
                else
                {
                    throw new Exception(msg);
                }
                AutomationHelper.ScrollToElement(driver, PVBancassuranceBTNXPath.CostCenter);

                AutomationHelper.SelectElement(driver, PVBancassuranceBTNXPath.CostCenter, param.CostCenterName);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
