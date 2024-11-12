using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;
using EPaymentVoucher.Locators.MyTaskList;
using EPaymentVoucher.Locators.Vendor;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.Vendor;
using EPaymentVoucher.Models.PaymentVoucher;
using EPaymentVoucher.Locators.PaymentVoucher;
using EPaymentVoucher.Models.General;
using EPaymentVoucher.Actions.PaymentVoucher;
using EPaymentVoucher.Actions.Vendor;

namespace EPaymentVoucher.Actions.MyTaskList
{
    public interface IMyTaskList
    {
        EdgeDriver VendorDatabaseHandleTaskAction(EdgeDriver driver, AppConfig cfg, TestPlan plan, List<VendorDatabaseTask> param, List<User> users, AutomationConfig allParam);
        EdgeDriver PaymentVoucherHandleTaskAction(EdgeDriver driver, AppConfig cfg, TestPlan plan, List<PVTask> param, List<User> users, AutomationConfig allParam);
    }

    public class MyTaskList(
        IHelperService utils,
        IPVReimbursement reimbursement,
        IPVAdvance advance,
        IVendor vendor,
        IPVSettlement settlement,
        IPVVendor pvVendor,
        IPVRDD rdd,
        IPVSalesForce salesForce,
        IPVAgencyCommission agencyCommission,
        IPVPolicyHolder policyHolder,
        IPVInternalBank internalBank,
        IPVSpecialCustomer specialCustomer,
        IPVBancassuranceBMI bancassuranceBMI,
        IPVBancassuranceBTN bancassuranceBTN
            ) : IMyTaskList
    {
        public Dictionary<int, string> _dictionaryResult = [];
        public Dictionary<int, string> _dictionaryResultPV = [];

        public EdgeDriver PaymentVoucherHandleTaskAction(EdgeDriver driver, AppConfig cfg, TestPlan plan, List<PVTask> param, List<User> users, AutomationConfig allParam)
        {
            _dictionaryResultPV.Clear();
            try
            {
                TestPlan newPlan = plan;
                _dictionaryResultPV = [];
                foreach (PVTask task in param)
                {
                    driver.Dispose();
                    driver = new();
                    User user = users.Where(x => x.Username == task.Actor).FirstOrDefault()!;
                    AutomationHelper.Login(driver, user, cfg);

                    var dataTask = allParam.TestPlans.Where(x => x.TestCaseId.Equals(plan.FromTestCaseId)).FirstOrDefault()!;
                    task.RequestNo = dataTask.RequestResult;
                    Console.WriteLine($"-- Task Approval. Test Case Id {newPlan.TestCaseId} - Sequence {task.Sequence} - Request No {task.RequestNo}");
                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.PVTask, task.Row);

                    string result = string.Empty;
                    if (task.Action.Equals(ApprovalConstant.Submit))
                    {
                        switch (plan.ModuleName)
                        {
                            case ModuleNameConstant.PVReimbursement:
                            case ModuleNameConstant.PVAdvance:
                            case ModuleNameConstant.PVVendor:
                            case ModuleNameConstant.PVPolicyHolder:
                            case ModuleNameConstant.PVInternalBank:
                            case ModuleNameConstant.PVSpecialCustomer:
                                result = PaymentVoucherTaskActionSubmit(driver, cfg, newPlan, task, allParam);
                                break;
                            default: break;
                        }
                    }
                    else
                    {
                        result = PaymentVoucherTaskAction(driver, cfg, newPlan, task, allParam);
                    }

                    if (result.Contains("Fail"))
                    {
                        _dictionaryResultPV.Add(task.Sequence, result);
                        break;
                    }
                }

                plan.Status = StatusConstant.Success;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";

                foreach (var item in _dictionaryResultPV.Values)
                {
                    if (item.Contains("Fail") || item.Contains("Fail."))
                    {
                        plan.Status = StatusConstant.Fail;
                        plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : One or more sequences have failed.";
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                plan.Status = StatusConstant.Fail;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : {ex.Message}";
            }
            finally
            {
                utils.WriteAutomationResult(cfg, plan);
                Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
                driver.Dispose();
            }
            return driver;
        }

        public string PaymentVoucherTaskActionSubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVTask param, AutomationConfig allParam)
        {
            string result = StatusConstant.Success;
            bool isAlert = false;
            try
            {
                SearchMTPV(driver, cfg, param);
                bool isExist = AutomationHelper.FindElements(driver, MTPaymentVoucherXPath.ListRowRequestNo);
                if (isExist)
                {
                    Thread.Sleep(5000);
                    AutomationHelper.ClickInteractableElementUsingJS(driver, MTPaymentVoucherXPath.ListRowRequestNo);

                    switch (plan.ModuleName)
                    {
                        case ModuleNameConstant.PVReimbursement:
                            PVReimbursementModels paramReimburs = allParam.PVReimbursements.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Submit)).FirstOrDefault()!;
                            if (paramReimburs != null)
                                reimbursement.HandleSubmit(driver, cfg, plan, paramReimburs);
                            break;
                        case ModuleNameConstant.PVAdvance:
                            PVAdvanceModels paramAdvance = allParam.PVAdvances.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Submit)).FirstOrDefault()!;
                            if (paramAdvance != null)
                                advance.HandleSubmit(driver, cfg, plan, paramAdvance);
                            break;
                        case ModuleNameConstant.PVVendor:
                            PVVendorModels parampvVendor = allParam.PVVendors.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Submit)).FirstOrDefault()!;
                            if (parampvVendor != null)
                                pvVendor.HandleSubmit(driver, cfg, plan, parampvVendor);
                            break;
                        case ModuleNameConstant.PVPolicyHolder:
                            PVPolicyHolderModels parampvpolicy = allParam.PVPolicyHolders.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Submit)).FirstOrDefault()!;
                            if (parampvpolicy != null)
                                policyHolder.HandleSubmit(driver, cfg, plan, parampvpolicy);
                            break;
                        case ModuleNameConstant.PVInternalBank:
                            PVInternalBankModels parampvinternalbank = allParam.PVInternalBanks.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Submit)).FirstOrDefault()!;
                            if (parampvinternalbank != null)
                                internalBank.HandleSubmit(driver, cfg, plan, parampvinternalbank);
                            break;
                        case ModuleNameConstant.PVSpecialCustomer:
                            PVSpecialCustomerModels parampvspecial = allParam.PVSpecialCustomers.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Submit)).FirstOrDefault()!;
                            if (parampvspecial != null)
                                specialCustomer.HandleSubmit(driver, cfg, plan, parampvspecial);
                            break;
                        default: break;
                    }

                    AutomationHelper.ScrollToElement(driver, PVApproval.Action);
                    Thread.Sleep(5000);
                    AutomationHelper.SelectElement(driver, PVApproval.Action, ApprovalConstant.Submit);
                    AutomationHelper.FillElementNonMandatory(driver, PVApproval.Notes, param.Notes);
                    AutomationHelper.ClickElement(driver, PVApproval.BtnSubmit);
                    Thread.Sleep(500);
                    string msg = AutomationHelper.GetAlertMessage(driver);
                    if (AutomationHelper.ValidateAlert(msg, "There is an error while ") || AutomationHelper.ValidateAlert(msg, "required"))
                    {
                        result = $"{msg}";
                        param.Result = result;
                        throw new Exception(result);
                    }
                    else
                    {
                        param.Result = StatusConstant.Success;
                    }
                }
                else
                {
                    result = $"Data not found";
                    param.Result = result;
                    throw new Exception(result);
                }
                isAlert = true;

            }
            catch (Exception ex)
            {
                result = $"Task {ApprovalConstant.Submit} {plan.ModuleName} is Fail. Error : {ex.Message}";
                param.Result = result;
            }
            finally
            {
                utils.WriteApproval(cfg, param, plan.ModuleName);
                Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
                if (isAlert)
                    AutomationHelper.HandleAlertJS(driver, true);
            }
            return result;
        }

        public string PaymentVoucherTaskAction(EdgeDriver driver, AppConfig cfg, TestPlan plan, PVTask param, AutomationConfig allParam)
        {
            string result = StatusConstant.Success;
            bool isAlert = false;
            try
            {
                SearchMTPV(driver, cfg, param);
                bool isExist = AutomationHelper.FindElements(driver, MTPaymentVoucherXPath.ListRowRequestNo);
                if (isExist)
                {
                    string status = driver.FindElement(By.XPath(MTPaymentVoucherXPath.RowStatus)).Text;
                    AutomationHelper.ClickElement(driver, MTPaymentVoucherXPath.ListRowRequestNo);
                    AutomationHelper.CheckElementExist(driver, PVApproval.Action);

                    if (param.Action.Equals(ApprovalConstant.Resubmit, StringComparison.OrdinalIgnoreCase))
                    {
                        switch (plan.ModuleName)
                        {
                            case ModuleNameConstant.PVReimbursement:
                                PVReimbursementModels paramReimburs = allParam.PVReimbursements.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (paramReimburs != null)
                                    reimbursement.HandleAddResubmit(driver, cfg, plan, paramReimburs, false);
                                break;
                            case ModuleNameConstant.PVAdvance:
                                PVAdvanceModels paramAdvance = allParam.PVAdvances.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (paramAdvance != null)
                                    advance.HandleAddResubmit(driver, cfg, plan, paramAdvance, false);
                                break;
                            case ModuleNameConstant.PVSettlement:
                                PVSettlementModels paramsettle = allParam.PVSettlements.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (paramsettle != null)
                                    settlement.HandleAddResubmit(driver, cfg, plan, paramsettle, false);
                                break;
                            case ModuleNameConstant.PVRDD:
                                PVSalesForceModels paramrdd = allParam.PVRDDs.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (paramrdd != null)
                                    rdd.HandleAddResubmit(driver, cfg, plan, paramrdd, false);
                                break;
                            case ModuleNameConstant.PVSalesForce:
                                PVSalesForceModels paramsalesforce = allParam.PVSalesForces.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (paramsalesforce != null)
                                    salesForce.HandleAddResubmit(driver, cfg, plan, paramsalesforce, false);
                                break;
                            case ModuleNameConstant.PVVendor:
                                PVVendorModels parampvVendor = allParam.PVVendors.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (parampvVendor != null)
                                    pvVendor.HandleAddResubmit(driver, cfg, plan, parampvVendor, false);
                                break;
                            case ModuleNameConstant.PVAgencyCommission:
                                PVAgencyCommissionModels parampvagencycom = allParam.PVAgencyCommissions.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (parampvagencycom != null)
                                    agencyCommission.HandleAddResubmit(driver, cfg, plan, parampvagencycom, false);
                                break;
                            case ModuleNameConstant.PVPolicyHolder:
                                PVPolicyHolderModels parampvpolicy = allParam.PVPolicyHolders.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (parampvpolicy != null)
                                    policyHolder.HandleAddResubmit(driver, cfg, plan, parampvpolicy, false);
                                break;
                            case ModuleNameConstant.PVInternalBank:
                                PVInternalBankModels parampvinternalbank = allParam.PVInternalBanks.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (parampvinternalbank != null)
                                    internalBank.HandleAddResubmit(driver, cfg, plan, parampvinternalbank, false);
                                break;
                            case ModuleNameConstant.PVSpecialCustomer:
                                PVSpecialCustomerModels parampvspecial = allParam.PVSpecialCustomers.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (parampvspecial != null)
                                    specialCustomer.HandleAddResubmit(driver, cfg, plan, parampvspecial, false);
                                break;
                            case ModuleNameConstant.PVBancassuranceBMI:
                                PVBancassuranceBMIModels bancaBMI = allParam.PVBancassuranceBMIs.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (bancaBMI != null)
                                    bancassuranceBMI.HandleAddResubmit(driver, cfg, plan, bancaBMI, false);
                                break;
                            case ModuleNameConstant.PVBancassuranceBTN:
                                PVBancassuranceBTNModels bancaBTN = allParam.PVBancassuranceBTNs.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
                                if (bancaBTN != null)
                                    bancassuranceBTN.HandleAddResubmit(driver, cfg, plan, bancaBTN, false);
                                break;
                            default: break;
                        }
                    }

                    AutomationHelper.ScrollToElement(driver, PVApproval.Action);
                    Thread.Sleep(5000);

                    if (param.Action.Equals(ApprovalConstant.Approve, StringComparison.OrdinalIgnoreCase) && status.Contains("finance", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            if (param.IsPayment)
                            {
                                AutomationHelper.ClickElement(driver, PVApproval.PaymentViaTrue);
                            }
                            else
                            {
                                AutomationHelper.ClickElement(driver, PVApproval.PaymentViaFalse);
                            }
                        }
                        catch (Exception) { Console.WriteLine($">>> Radio Button Hidden"); }
                    }

                    if (param.Action.Equals(ApprovalConstant.Return, StringComparison.OrdinalIgnoreCase))
                    {
                        AutomationHelper.SelectElement(driver, PVApproval.Action, ApprovalConstant.Return);
                        if (string.IsNullOrEmpty(param.ReturnTo))
                        {
                            result = $"Return To cannot be empty";
                            param.Result = result;
                            throw new Exception(result);
                        }
                        else
                        {
                            AutomationHelper.SelectElement(driver, PVApproval.ReturnTo, param.ReturnTo);
                        }
                    }
                    else
                    {
                        AutomationHelper.SelectElement(driver, PVApproval.Action, param.Action);
                    }

                    AutomationHelper.FillElementNonMandatory(driver, PVApproval.Notes, param.Notes);

                    if (param.Action.Equals(ApprovalConstant.Reject, StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(param.Notes))
                    {
                        result = $"Notes cannot be empty";
                        param.Result = result;
                        throw new Exception(result);
                    }
                    else
                    {
                        AutomationHelper.ScrollToElement(driver, PVApproval.BtnSubmit);
                        AutomationHelper.ClickElement(driver, PVApproval.BtnSubmit);
                        Thread.Sleep(500);
                        string msg = AutomationHelper.GetAlertMessage(driver);
                        if (AutomationHelper.ValidateAlert(msg, "There is an error while ") || AutomationHelper.ValidateAlert(msg, "required"))
                        {
                            result = $"{msg}";
                            param.Result = result;
                            throw new Exception(result);
                        }
                        else
                        {
                            param.Result = StatusConstant.Success;
                        }
                    }
                }
                else
                {
                    result = $"Data not found";
                    param.Result = result;
                    throw new Exception(result);
                }
                isAlert = true;

            }
            catch (Exception ex)
            {
                result = $"Task {param.Action} {plan.ModuleName} is Fail. Error : {ex.Message}";
                param.Result = result;
            }
            finally
            {
                utils.WriteApproval(cfg, param, plan.ModuleName);
                Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
                if (isAlert)
                    AutomationHelper.HandleAlertJS(driver, true);
            }
            return result;
        }

        private static void SearchMTPV(EdgeDriver driver, AppConfig cfg, PVTask param)
        {
            try
            {
                OpenPaymentVoucher(driver, cfg.Url);
                Thread.Sleep(5000);
                AutomationHelper.CheckElementExist(driver, MTPaymentVoucherXPath.BtnSearch);

                AutomationHelper.FillElementNonMandatory(driver, MTPaymentVoucherXPath.RequestNo, param.RequestNo);

                AutomationHelper.ClickElement(driver, MTPaymentVoucherXPath.BtnSearch);
                Thread.Sleep(5000);
            }
            catch (Exception ex)
            {
                throw new Exception($"Search Task Payment Voucher is Fail : {ex.Message}");
            }
        }

        public EdgeDriver VendorDatabaseHandleTaskAction(EdgeDriver driver, AppConfig cfg, TestPlan plan, List<VendorDatabaseTask> param, List<User> users, AutomationConfig allParam)
        {
            _dictionaryResult.Clear();
            try
            {
                string userLogin = plan.UserLogin;
                TestPlan newPlan = plan;
                _dictionaryResult = [];

                foreach (VendorDatabaseTask task in param)
                {
                    driver.Dispose();
                    driver = new();
                    User user = users.Where(x => x.Username == task.Actor).FirstOrDefault()!;
                    AutomationHelper.Login(driver, user, cfg);
                    VendorDatabase vendorParam = task.Action.Equals(ApprovalConstant.Resubmit) ? allParam.VendorDatabases.Where(x => x.TestCaseId.Equals(task.TestCaseId) && x.DataFor.Equals(DataConstant.Resubmit)).FirstOrDefault()! : new();
                    TestPlan dataSearch = allParam.TestPlans.Where(x => x.TestCaseId.Equals(plan.FromTestCaseId)).FirstOrDefault()!;
                    task.ParamSearch.VendorName = dataSearch.RequestResult;
                    if (vendorParam != null && dataSearch != null)
                    {
                        vendorParam.SearchVendor.VendorName = dataSearch.RequestResult;
                    }
                    plan.TestData = PlanHelper.CreateAddress(SheetConstant.VendorDatabaseTask, task.Row);
                    string result = VendorDatabaseTaskAction(driver, cfg, newPlan, task, vendorParam!);
                    if (result.Contains("Fail"))
                    {
                        _dictionaryResult.Add(task.Sequence, result);
                        driver.Dispose();
                        break;
                    }
                }

                plan.Status = StatusConstant.Success;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";

                foreach (var item in _dictionaryResult.Values)
                {
                    if (item.Contains("Fail") || item.Contains("Fail."))
                    {
                        plan.Status = StatusConstant.Fail;
                        plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : One or more sequences have failed.";
                        driver.Dispose();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                plan.Status = StatusConstant.Fail;
                plan.Remarks = $"Task {plan.TestCase} {ModuleNameConstant.VendorDatabase} is Fail. Error : {ex.Message}";
            }
            finally
            {
                utils.WriteAutomationResult(cfg, plan);
                Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
                driver.Dispose();
            }
            return driver;
        }

        public string VendorDatabaseTaskAction(EdgeDriver driver, AppConfig cfg, TestPlan plan, VendorDatabaseTask param, VendorDatabase vendorParam)
        {
            string result = StatusConstant.Success;
            try
            {
                VendorDatabaseSearchTask searchParam = param.ParamSearch;

                SearchMTVendor(driver, cfg, searchParam);
                string xpathReqNo = param.Action.Equals(ApprovalConstant.Resubmit, StringComparison.OrdinalIgnoreCase) || param.Action.Equals(ApprovalConstant.Cancel, StringComparison.OrdinalIgnoreCase)
                    ? MTVendorXPath.RowsRequestNoResubmit
                    : MTVendorXPath.RowsRequestNo;
                AutomationHelper.CheckElementExist(driver, xpathReqNo);
                var listRows = driver.FindElements(By.XPath(xpathReqNo));
                if (listRows.Count > 0)
                {
                    if (param.Action.Equals(ApprovalConstant.Approve, StringComparison.OrdinalIgnoreCase))
                    {
                        listRows.First().Click();
                        AutomationHelper.CheckElementExist(driver, MTVendorXPath.RowRole);

                        string lastRole = driver.FindElement(By.XPath(MTVendorXPath.RowRole)).Text.Trim();
                        string lastAction = driver.FindElement(By.XPath(MTVendorXPath.RowAction)).Text.Trim();

                        if (lastRole.Equals("compliance", StringComparison.OrdinalIgnoreCase) && lastAction.Equals(ApprovalConstant.Approve, StringComparison.OrdinalIgnoreCase))
                        {
                            AutomationHelper.ClickElement(driver, MTVendorXPath.TabProfile);
                            Thread.Sleep(500);

                            var vendorCodeField = driver.FindElement(By.XPath(MTVendorXPath.FieldVendorCode));
                            if (string.IsNullOrEmpty(vendorCodeField.Text))
                            {
                                if (string.IsNullOrEmpty(param.NewVendorCode))
                                {
                                    throw new Exception($"{param.Action} {plan.ModuleName} is Fail. Msg : New Vendor Code cannot be empty at this sequence {param.Sequence}.");
                                }
                                vendorCodeField.SendKeys(param.NewVendorCode);
                            }

                            AutomationHelper.ClickElement(driver, MTVendorXPath.TabApproval);
                        }

                        AutomationHelper.SelectElement(driver, MTVendorXPath.Action, param.Action);

                        AutomationHelper.FillElementNonMandatory(driver, MTVendorXPath.Notes, param.Notes);

                        AutomationHelper.ClickElement(driver, MTVendorXPath.BtnSubmitAction);
                        string err = AutomationHelper.CheckAlertErrorVendorDatabase(driver);
                        if (!string.IsNullOrEmpty(err))
                        {
                            throw new Exception(err);
                        }
                        Thread.Sleep(2000);
                        param.Result = StatusConstant.Success;
                    }
                    else
                    {
                        listRows.First().Click();
                        IWebElement btnSubmit;

                        if (param.Action.Equals(ApprovalConstant.Resubmit, StringComparison.OrdinalIgnoreCase) || param.Action.Equals(ApprovalConstant.Cancel, StringComparison.OrdinalIgnoreCase))
                        {
                            if (param.Action.Equals(ApprovalConstant.Cancel, StringComparison.OrdinalIgnoreCase))
                            {
                                AutomationHelper.ScrollToElement(driver, VendorListXPath.FormBtnNext);
                                AutomationHelper.ClickElement(driver, VendorListXPath.FormBtnNext);
                                AutomationHelper.ScrollToElement(driver, VendorListXPath.FormDetailBtnNext);
                                AutomationHelper.ClickElement(driver, VendorListXPath.FormDetailBtnNext);
                            }
                            else
                            {
                                if (vendorParam != null)
                                    vendor.HandleAddEditResubmit(driver, cfg, plan, vendorParam, true, true);
                                AutomationHelper.ScrollToElement(driver, VendorListXPath.FormDetailBtnNext);
                                AutomationHelper.ClickElement(driver, VendorListXPath.FormDetailBtnNext);
                            }
                            btnSubmit = driver.FindElement(By.XPath(MTVendorXPath.BtnSubmitActionResubmit));
                        }
                        else
                        {
                            btnSubmit = driver.FindElement(By.XPath(MTVendorXPath.BtnSubmitAction));
                        }


                        AutomationHelper.SelectElement(driver, MTVendorXPath.Action, param.Action);

                        if (param.Action.Equals(ApprovalConstant.Return, StringComparison.OrdinalIgnoreCase))
                        {
                            AutomationHelper.SelectElement(driver, MTVendorXPath.ReturnTo, param.ReturnTo);
                        }

                        AutomationHelper.FillElementNonMandatory(driver, MTVendorXPath.Notes, param.Notes);

                        if (param.Action.Equals(ApprovalConstant.Reject, StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(param.Notes))
                        {
                            result = $"{param.Action} {plan.ModuleName} is Fail. Msg : Notes cannot be empty";
                            param.Result = result;
                        }
                        else
                        {
                            AutomationHelper.ScrollToElement(driver, btnSubmit);
                            btnSubmit.Click();
                            string err = AutomationHelper.CheckAlertErrorVendorDatabase(driver);
                            if (!string.IsNullOrEmpty(err))
                            {
                                throw new Exception(err);
                            }
                            Thread.Sleep(2000);
                            param.Result = StatusConstant.Success;
                        }
                    }
                }
                else
                {
                    result = $"{param.Action} {plan.ModuleName} is Fail. Notes : No records of data were found for the Automation Task {param.Action} {plan.ModuleName}.";
                    param.Result = result;
                }

            }
            catch (Exception ex)
            {
                result = $"{param.Action} {plan.ModuleName} is Fail. Error : {ex.Message}";
                param.Result = result;
            }
            finally
            {
                utils.WriteApproval(cfg, param, ModuleNameConstant.VendorDatabase);
                Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
            }

            return result;
        }

        private static void SearchMTVendor(EdgeDriver driver, AppConfig cfg, VendorDatabaseSearchTask param)
        {
            try
            {
                OpenMyTaskList(driver, cfg.Url);
                AutomationHelper.ClickElement(driver, MTVendorXPath.TabMTVD);
                AutomationHelper.CheckElementExist(driver, MTVendorXPath.MTVDVBtnSearch);

                AutomationHelper.FillElementNonMandatory(driver, MTVendorXPath.MTVDVendorCode, param.VendorCode);
                AutomationHelper.FillElementNonMandatory(driver, MTVendorXPath.MTVDVendorName, param.VendorName);
                AutomationHelper.FillElementNonMandatory(driver, MTVendorXPath.MTVDBusinessName, param.BusinessName);

                AutomationHelper.ClickElement(driver, MTVendorXPath.MTVDVBtnSearch);
            }
            catch (Exception ex)
            {
                throw new Exception($"Search Task Vendor Database is Fail : {ex.Message}");
            }
        }

        private static void OpenMyTaskList(EdgeDriver driver, string url)
        {
            try
            {
                driver.Navigate().GoToUrl($"{url}/Tasks/MyTaskList");
            }
            catch (Exception ex)
            {
                throw new Exception($"Open Tab Task Vendor Database is Fail : {ex.Message}");
            }
        }

        private static void OpenPaymentVoucher(EdgeDriver driver, string url)
        {
            try
            {
                driver.Navigate().GoToUrl($"{url}{MTPaymentVoucherXPath.TabMTPV}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Open Tab Task Payment Voucher is Fail : {ex.Message}");
            }
        }
    }
}
