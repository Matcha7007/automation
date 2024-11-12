using OpenQA.Selenium.Edge;
using EPaymentVoucher.Models;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.Vendor;
using EPaymentVoucher.Actions.MyTaskList;
using EPaymentVoucher.Actions.PaymentVoucher;
using EPaymentVoucher.Actions.Vendor;
using EPaymentVoucher.Actions.BankConfirmation;

namespace EPaymentVoucher.Utilities
{
	public interface IUnitTestHandler
	{
		void RunUnitTest(AppConfig cfg);
	}
	public class UnitTestHandler(
			IHelperService utils,
			IVendor vendor,
			IPaymentVoucher pv,
			IMyTaskList task,
			IBankConfirmation bankConfirmation
			) : IUnitTestHandler
	{
		public void RunUnitTest(AppConfig cfg)
		{
			try
			{
				AutomationConfig automationConfig = utils.ReadAutomationConfig(cfg);

				foreach (var plan in automationConfig.TestPlans)
				{
					if (!string.IsNullOrEmpty(plan.Status))
						continue;
					Console.WriteLine("------------------------------------------------------");
					Console.WriteLine($"Running Unit Test {plan.TestCase} - {plan.ModuleName}");
					Console.WriteLine("------------------------------------------------------");
					EdgeDriver driver = new();
					driver.Manage().Window.FullScreen();
					
					try
					{
						SwitchCaseTest(cfg, driver, plan, automationConfig.Users, automationConfig);
						Console.WriteLine("------------------------------------------------------");
						Console.WriteLine($"End of Unit Test {plan.TestCase} - {plan.ModuleName}.");
						Console.WriteLine("------------------------------------------------------");
					}
					catch (Exception ex)
					{
						Console.WriteLine("------------------------------------------------------");
						Console.WriteLine($"End of Unit Test {plan.TestCase} - {plan.ModuleName} is Fail. Msg : {ex.Message}");
						Console.WriteLine("------------------------------------------------------");
					}

					driver.Dispose();
				}
			}
			catch (Exception ex) 
			{
				Console.WriteLine("------------------------------------------------------");
				Console.WriteLine($"Run Unit Test Failed : {ex.Message}");
				Console.WriteLine("------------------------------------------------------");
			}
		}
	
		private void SwitchCaseTest(AppConfig cfg, EdgeDriver driver, TestPlan plan, List<User> users, AutomationConfig param)
		{
			try
			{
				User user = users.Where(x => x.Username == plan.UserLogin).FirstOrDefault()!;
				if (plan.TestCase == TestCaseConstant.Add || plan.TestCase == TestCaseConstant.Edit || plan.ModuleName == ModuleNameConstant.RejectBankConfirmation)
					AutomationHelper.Login(driver, user, cfg);

				Console.WriteLine($"plan id : {plan.TestCaseId} - module name : {plan.ModuleName}");

				switch (plan.ModuleName)
				{
					case ModuleNameConstant.RejectBankConfirmation:
						bankConfirmation.HandleTestCase(driver, cfg, plan, param);
						break;
					case ModuleNameConstant.VendorDatabase:
						switch (plan.TestCase)
						{
							case TestCaseConstant.Add:
							case TestCaseConstant.Edit:
								vendor.HandleTestCase(driver, cfg, plan, param, users);
								break;
							case TestCaseConstant.Approval:
								List<VendorDatabaseTask> taskData = [.. param.VendorDatabaseTasks.Where(x => x.TestCaseId == plan.TestCaseId).OrderBy(x => x.Sequence)];
								if (taskData.Count > 0)
								{
									plan.TestData = PlanHelper.CreateAddress(SheetConstant.VendorDatabaseTask, taskData.FirstOrDefault()!.Row);
									driver = task.VendorDatabaseHandleTaskAction(driver, cfg, plan, taskData, users, param);
								}
								break;
							default: break;
						}
						break;
					case ModuleNameConstant.PVReimbursement:
					case ModuleNameConstant.PVAdvance:
					case ModuleNameConstant.PVSettlement:
					case ModuleNameConstant.PVRDD:
					case ModuleNameConstant.PVSalesForce:
					case ModuleNameConstant.PVVendor:
					case ModuleNameConstant.PVPolicyHolder:
					case ModuleNameConstant.PVInternalBank:
					case ModuleNameConstant.PVSpecialCustomer:
					case ModuleNameConstant.PVAgencyCommission:
					case ModuleNameConstant.PVBancassuranceBMI:
					case ModuleNameConstant.PVBancassuranceBTN:
						Console.WriteLine("Masuk handle pv");
						pv.HandleTestCase(driver, cfg, plan, param, users);
						break;
					default:
						break;
				}
			}
			catch (Exception ex) { throw new Exception(ex.Message); }
			finally { driver.Dispose(); }
		}
	}
}
