using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models;
using OpenQA.Selenium.Edge;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Utilities;
using EPaymentVoucher.Actions.MyTaskList;

namespace EPaymentVoucher.Actions.PaymentVoucher
{
    public interface IPaymentVoucher
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param, List<User> users);
    }

    public class PaymentVoucher(
        IPVReimbursement reimbursement,
        IPVAdvance advance,
        IMyTaskList task,
        IPVSettlement settlement,
        IPVVendor vendor,
        IPVRDD rdd,
        IPVSalesForce salesForce,
        IPVAgencyCommission agencyCommission,
        IPVPolicyHolder policyHolder,
        IPVInternalBank internalBank,
        IPVSpecialCustomer specialCustomer,
        IPVBancassuranceBMI bancassuranceBMI,
        IPVBancassuranceBTN bancassuranceBTN
            ) : IPaymentVoucher
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param, List<User> users)
        {
            try
            {
                List<PVTask> pvTaskData = param.PVTasks.Where(x => x.TestCaseId == plan.TestCaseId).OrderBy(x => x.Sequence).ToList();
                switch (plan.ModuleName)
                {
                    case ModuleNameConstant.PVReimbursement:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                reimbursement.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVAdvance:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                advance.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVSettlement:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                settlement.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVRDD:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                rdd.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVSalesForce:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                salesForce.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVVendor:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                vendor.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVAgencyCommission:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                agencyCommission.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVPolicyHolder:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                policyHolder.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVInternalBank:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                internalBank.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVSpecialCustomer:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                specialCustomer.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVBancassuranceBMI:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                bancassuranceBMI.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    case ModuleNameConstant.PVBancassuranceBTN:
                        switch (plan.TestCase)
                        {
                            case TestCaseConstant.Add:
                                bancassuranceBTN.HandleTestCase(driver, cfg, plan, param);
                                break;
                            case TestCaseConstant.Approval:
                                if (pvTaskData.Count > 0)
                                {
                                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, plan.Row);
                                    driver = task.PaymentVoucherHandleTaskAction(driver, cfg, plan, pvTaskData, users, param);
                                }
                                break;
                            default: break;
                        }
                        break;
                    default: break;
                }
            }
            catch (Exception ex) { throw new Exception($"Handle Test Case PV - {ex.Message}"); }
            finally { driver.Dispose(); }
        }
    }
}
