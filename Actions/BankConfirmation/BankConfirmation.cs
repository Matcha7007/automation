using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;
using OpenQA.Selenium.Edge;
using EPaymentVoucher.Models.BankConfirmation;
using EPaymentVoucher.Locators.BankConfirmation;

namespace EPaymentVoucher.Actions.BankConfirmation
{
	public interface IBankConfirmation
	{
		void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param);
	}
	public class BankConfirmation(IHelperService utils) : IBankConfirmation
	{
		public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param)
		{
			try
			{
				switch (plan.ModuleName)
				{
					case ModuleNameConstant.RejectBankConfirmation:
						BankConfirmationModels pv = param.RejectBankConfirmations.Where(x => x.TestCaseId == plan.TestCaseId).FirstOrDefault()!;
						plan.TestData = PlanHelper.CreateAddress(SheetConstant.RejectBankConfirmation, pv.Row);
						HandleReject(driver, cfg, plan, pv);
						break;
					default: break;
				}
			}
			catch (Exception ex) { throw new Exception($"Handle Test Case Bank Confirmation - {ex.Message}"); }
			finally { driver.Dispose(); }
		}

		public void HandleReject(EdgeDriver driver, AppConfig cfg, TestPlan plan, BankConfirmationModels param, bool isNew = true)
		{
			bool isAlert = false;
			plan.RequestResult = param.RequestNo;
			try
			{
				AutomationHelper.OpenPage(driver, $"{cfg.Url}{BankConfirmationXPath.UrlBankConfirmationList}");

				HandleSearch(driver, param);

				HandleAction(driver, param);

				string msg = AutomationHelper.GetAlertMessage(driver);
				if (AutomationHelper.ValidateAlert(msg, "success"))
				{
					plan.Status = StatusConstant.Success;
					plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
				}
				else
				{
					throw new Exception(msg);
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
					utils.WriteAutomationResult(cfg, plan);
					Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
					if (isAlert)
						AutomationHelper.HandleAlertJS(driver, true);
				}
			}
		}

		public static void HandleAction(EdgeDriver driver, BankConfirmationModels param)
		{
			try
			{
				AutomationHelper.ClickElement(driver, BankConfirmationXPath.RequestPaymentVoucherList);
				AutomationHelper.ClickElement(driver, BankConfirmationXPath.DetailTransactionAll);
				AutomationHelper.ClickElement(driver, BankConfirmationXPath.BtnSelectDetailTransaction);

				AutomationHelper.ScrollToElement(driver, BankConfirmationXPath.BtnReturn);
				AutomationHelper.DatePicker(driver, BankConfirmationXPath.TransferDate, param.TransferDate);
				AutomationHelper.SelectElement(driver, BankConfirmationXPath.BankSource, param.BankSource);
				AutomationHelper.FillElement(driver, BankConfirmationXPath.Rate, param.Rate, true);
				AutomationHelper.ClickElement(driver, BankConfirmationXPath.BtnCalculate);
				AutomationHelper.ClickElement(driver, BankConfirmationXPath.BtnReturn);
			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		public static void HandleSearch(EdgeDriver driver, BankConfirmationModels param)
		{
			try
			{
				AutomationHelper.FillElement(driver, BankConfirmationXPath.RequestNo, param.RequestNo);
				AutomationHelper.ClickElement(driver, BankConfirmationXPath.BtnSearch);

				try
				{
					AutomationHelper.ClickElement(driver, BankConfirmationXPath.RequestPaymentVoucherList);
				}
				catch
				{
					throw new Exception("Data not found");
				}
			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}
	}
}
