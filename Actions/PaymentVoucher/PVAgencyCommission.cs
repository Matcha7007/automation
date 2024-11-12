using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;

using OpenQA.Selenium.Edge;
using EPaymentVoucher.Locators.PaymentVoucher;
using EPaymentVoucher.Models.General;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public interface IPVAgencyCommission
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
        void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVAgencyCommissionModels param, bool isNew = true);
    }
    public class PVAgencyCommission(IHelperService utils) : IPVAgencyCommission
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        PVAgencyCommissionModels pv = param.PVAgencyCommissions.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVAgencyCommission, pv.Row);
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

        public void HandleAddResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVAgencyCommissionModels param, bool isNew = true)
        {
            PVBaseModels pVBase = new() { Title = param.Title, PaymentPurpose = param.PaymentPurpose, PaymentType = PVPaymentTypeConst.AgencyCommission, AttachmentDescription = param.AttachmentDescription, AttachmentPath = param.AttachmentPath, Remarks = param.Remarks };
            bool isAlert = false;
            try
            {
                if (isNew)
                {
                    AutomationHelper.OpenPage(driver, $"{cfg.Url}{PVListXPath.UrlCreatePaymentList}");

                    PVHelper.FillFirstPage(driver, pVBase);
                }

                HandleUpload(driver, param.PaymentPurpose, param, isNew);

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

        public static void HandleUpload(EdgeDriver driver, string purpose, PVAgencyCommissionModels param, bool isNew)
        {
            try
            {
                if (purpose.Equals(PaymentPurposeAgencyCommissionConst.GAAllowancePT, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.ScrollToElement(driver, ACGAAllowancePTXPath.BtnFile);
                    AutomationHelper.UploadFile(driver, ACGAAllowancePTXPath.BtnFile, param.GAAllowancePTPath);
                }
                else if (purpose.Equals(PaymentPurposeAgencyCommissionConst.Financing, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.ScrollToElement(driver, ACFinancingXPath.BtnFile);
                    AutomationHelper.UploadFile(driver, ACFinancingXPath.BtnFile, param.FinancingPath);
                }
                else if (purpose.Equals(PaymentPurposeAgencyCommissionConst.MGI, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.ScrollToElement(driver, ACMGIXPath.BtnFile);
                    AutomationHelper.UploadFile(driver, ACMGIXPath.BtnFile, param.MGIPath);
                }
                else if (purpose.Equals(PaymentPurposeAgencyCommissionConst.CommissionMidEndMonth, StringComparison.OrdinalIgnoreCase))
                {
                    AutomationHelper.ScrollToElement(driver, ACCommisionMEMonthXPath.BtnTaxCalculation);
                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnTaxCalculation, param.TaxCalculationPath);
                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnFirstYearCommission, param.FirstYearCRPath);
                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnRenewalYearCommission, param.RenewalYearCRPath);
                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnReportRecOverride, param.ReportRecOverridePath);
                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnADGenDetail, param.ADGenDetailPath);
                    AutomationHelper.UploadFileNonMandatory(driver, ACCommisionMEMonthXPath.BtnProducerBonusReport, param.ProducerBonusReportPath);

                    AutomationHelper.ScrollToElement(driver, ACCommisionMEMonthXPath.BtnActivationAgencyLeaderBonusReport);
                    AutomationHelper.UploadFileNonMandatory(driver, ACCommisionMEMonthXPath.BtnActivationAgencyLeaderBonusReport, param.AALBonusReportPath);
                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnSummaryOverride, param.SummaryOverridePath);
                    AutomationHelper.UploadFileNonMandatory(driver, ACCommisionMEMonthXPath.BtnGAAllowancePersonal, param.GAAllowancePersonalPath);
                    AutomationHelper.UploadFileNonMandatory(driver, ACCommisionMEMonthXPath.BtnSummaryAgentPaid, param.SummaryAgentPaidPath);

                    AutomationHelper.UploadFile(driver, ACCommisionMEMonthXPath.BtnCFGManualPayment, param.CFGManualPaymentPath);
                }

                AutomationHelper.ClickElement(driver, ACSummaryXPath.BaseBtnUploadFile);
                string msg = AutomationHelper.GetAlertMessage(driver);
                if (AutomationHelper.ValidateAlert(msg, "success"))
                {
                    AutomationHelper.HandleAlertJS(driver, true);
                }
                else
                {
                    throw new Exception(msg);
                }
                AutomationHelper.ScrollToElement(driver, ACSummaryXPath.CostCenter);
                AutomationHelper.SelectElement(driver, ACSummaryXPath.CostCenter, param.CostCenterName);
            }
            catch (Exception ex) { throw new Exception(ex.Message); }
        }
    }
}
