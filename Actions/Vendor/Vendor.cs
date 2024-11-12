using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;
using EPaymentVoucher.Models.Vendor;
using EPaymentVoucher.Locators.Vendor;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models.General;

namespace EPaymentVoucher.Actions.Vendor
{
    public interface IVendor
    {
        void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param, List<User> users);
        void HandleAddEditResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, VendorDatabase param, bool isEdit, bool isResubmit = false);
    }

    public class Vendor(IHelperService utils) : IVendor
    {
        public void HandleTestCase(EdgeDriver driver, AppConfig cfg, TestPlan plan, AutomationConfig param, List<User> users)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Add:
                        VendorDatabase add = param.VendorDatabases.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Add)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.VendorDatabase, add.Row);
                        plan.RequestResult = add.VendorName;
                        HandleAddEditResubmit(driver, cfg, plan, add, plan.TestCase.Equals(TestCaseConstant.Edit));
                        break;
                    case TestCaseConstant.Edit:
                        VendorDatabase vendorParam = param.VendorDatabases.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataConstant.Edit)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.VendorDatabase, vendorParam.Row);
                        if (plan.TestCase.Equals(TestCaseConstant.Edit))
                        {
                            TestPlan dataSearch = param.TestPlans.Where(x => x.TestCaseId.Equals(plan.FromTestCaseId)).FirstOrDefault()!;
                            vendorParam.SearchVendor.VendorName = dataSearch.RequestResult;
                        }
                        plan.RequestResult = vendorParam.VendorName;
                        HandleAddEditResubmit(driver, cfg, plan, vendorParam, plan.TestCase.Equals(TestCaseConstant.Edit));
                        break;
                    default: break;
                }
            }
            finally
            {
                driver.Dispose();
            }
        }

        public void HandleAddEditResubmit(EdgeDriver driver, AppConfig cfg, TestPlan plan, VendorDatabase param, bool isEdit, bool isResubmit = false)
        {
            try
            {
                bool isSuccess = true;

                #region Condition Add / Edit
                if (!isResubmit)
                {
                    if (isEdit)
                    {
                        SearchVendor(driver, cfg, plan, param);
                        var lists = driver.FindElements(By.XPath(VendorListXPath.BtnEdit));

                        if (lists.Count > 0)
                        {
                            AutomationHelper.ClickInteractableElement(driver, VendorListXPath.BtnEdit);
                        }
                        else
                        {
                            plan.Status = StatusConstant.Success;
                            plan.Remarks = $"{plan.TestCase} {plan.ModuleName} : Not found the data.";
                            isSuccess = false;
                        }
                    }
                    else
                    {
                        driver.Navigate().GoToUrl($"{cfg.Url}/Vendor/Create");
                    }
                }
                #endregion

                if (isEdit && isSuccess || !isEdit)
                {
                    #region Profile
                    VendorProfile(driver, cfg, plan, param, isEdit, isResubmit);
                    #endregion

                    #region Detail
                    if (param.BankAccounts.Count > 0 && param.BusinessNames.Count > 0)
                    {
                        VendorDetail(driver, param, isEdit, isResubmit);
                    }
                    else
                    {
                        throw new Exception($"Parameter at sheet {SheetConstant.VendorDatabaseBankAccount} or {SheetConstant.VendorDatabaseBusinessName} can't be empty for Test Case Id : {param.TestCaseId}.");
                    }
                    #endregion

                    if (!isResubmit)
                    {
                        AutomationHelper.ClickElement(driver, VendorListXPath.FormDetailBtnSave);
                        string err = AutomationHelper.CheckAlertErrorVendorDatabase(driver);
                        if (!string.IsNullOrEmpty(err))
                        {
                            throw new Exception(err);
                        }
                    }
                }
                plan.Status = StatusConstant.Success;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
                plan.RequestResult = param.VendorName;
            }
            catch (Exception ex)
            {
                plan.Status = StatusConstant.Fail;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : {ex.Message}";
            }
            finally
            {
                if (!isResubmit)
                {
                    utils.WriteAutomationResult(cfg, plan);
                    Thread.Sleep(5000);
                }
            }
        }

        private static void VendorDetail(EdgeDriver driver, VendorDatabase param, bool isEdit, bool isResubmit)
        {
            List<VendorContactPersonParam> contacts = param.ContactPersons;
            List<VendorBankAccountParam> bankAccounts = param.BankAccounts;
            List<VendorBusinessNameParam> businessNames = param.BusinessNames;

            AutomationHelper.ScrollToElement(driver, VendorListXPath.FormDBtnAddBankAccount);

            #region Contact Person
            if (isEdit)
            {
                var listContacts = driver.FindElements(By.XPath(VendorListXPath.FormDBtnDeleteContactPerson));
                foreach (var contact in listContacts)
                {
                    contact.Click();
                    Thread.Sleep(1000);
                }
            }
            foreach (VendorContactPersonParam contact in contacts)
            {
                AutomationHelper.ClickInteractableElementUsingJS(driver, VendorListXPath.FormDBtnAddContactPerson);
                AutomationHelper.CheckElementExist(driver, VendorListXPath.FormDCPName);

                AutomationHelper.FillElement(driver, VendorListXPath.FormDCPName, contact.ContactPersonName);
                FillPhone(driver, VendorListXPath.FormDCPPhone1_Front, VendorListXPath.FormDCPPhone1_Main, VendorListXPath.FormDCPPhone1_Back, contact.ContactPersonPhone1!);
                if (!string.IsNullOrEmpty(contact.ContactPersonPhone2))
                    FillPhone(driver, VendorListXPath.FormDCPPhone2_Front, VendorListXPath.FormDCPPhone2_Main, VendorListXPath.FormDCPPhone2_Back, contact.ContactPersonPhone2);

                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.FormDCPEmail, contact.ContactPersonEmail);

                AutomationHelper.ClickElement(driver, VendorListXPath.FormDCPBtnAdd);
                Thread.Sleep(1000);
            }
            #endregion

            #region Bank Account
            if (isEdit)
            {

                var listBankAccounts = driver.FindElements(By.XPath(VendorListXPath.FormDBtnDeleteBankAccount));
                foreach (var bankAccount in listBankAccounts)
                {
                    bankAccount.Click();
                }
            }
            foreach (VendorBankAccountParam bankAccount in bankAccounts)
            {
                AutomationHelper.ClickInteractableElementUsingJS(driver, VendorListXPath.FormDBtnAddBankAccount);
                AutomationHelper.CheckElementExist(driver, VendorListXPath.FormDBACountry);

                AutomationHelper.SelectElement(driver, VendorListXPath.FormDBACountry, bankAccount.BankAccountCountry);

                AutomationHelper.FillElement(driver, VendorListXPath.FormDBAAliasName, bankAccount.BankAccountAliasName);

                if (param.Country!.ToLower() == "indonesia")
                {
                    AutomationHelper.SelectElement(driver, VendorListXPath.FormDBABankName, bankAccount.BankAccountBankName);
                }
                else
                {
                    AutomationHelper.FillElement(driver, VendorListXPath.FormDBABankNameInput, bankAccount.BankAccountBankName);
                }

                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.FormDBABranch, bankAccount.BankAccountBranch);

                AutomationHelper.FillElement(driver, VendorListXPath.FormDBAAccountNo, bankAccount.BankAccountNo);
                AutomationHelper.FillElement(driver, VendorListXPath.FormDBAAccountName, bankAccount.BankAccountName);

                AutomationHelper.SelectElement(driver, VendorListXPath.FormDBACurrency, bankAccount.BankAccountCurrency);

                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.FormDBASwiftCode, bankAccount.BankAccountSwiftCode);
                if (bankAccount.BankAccountIsDefault)
                    AutomationHelper.ClickElement(driver, VendorListXPath.FormDBADefault);

                AutomationHelper.UploadFile(driver, VendorListXPath.FormDBABtnUpload, bankAccount.BankAccountEvidencePath);

                AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormDBABtnAdd);
                Thread.Sleep(1000);
            }
            #endregion

            AutomationHelper.ScrollToElement(driver, VendorListXPath.FormDetailBtnSave);

            #region Business Name
            if (isEdit)
            {
                var listBusinessNames = driver.FindElements(By.XPath(VendorListXPath.FormDBtnDeleteVendorBusiness));
                foreach (var businessName in listBusinessNames)
                {
                    businessName.Click();
                    Thread.Sleep(500);
                }
            }
            foreach (VendorBusinessNameParam businessName in businessNames)
            {
                AutomationHelper.ClickInteractableElementUsingJS(driver, VendorListXPath.FormDBtnAddVendorBusiness);
                AutomationHelper.CheckElementExist(driver, VendorListXPath.FormDVBCategory);

                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.FormDVBCategory, businessName.BusinessNameCategory);
                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.FormDVBName, businessName.BusinessName);
                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.FormDVBType, businessName.BusinessNameType);

                //search
                AutomationHelper.ClickElement(driver, VendorListXPath.FormDVBBtnSearch);
                Thread.Sleep(1000);
                AutomationHelper.ClickElement(driver, VendorListXPath.FormDVBTableTrChecked);
                AutomationHelper.ClickElement(driver, VendorListXPath.FormDVBBtnSelect);
                Thread.Sleep(1000);
            }
            #endregion
        }

        private static void VendorProfile(EdgeDriver driver, AppConfig cfg, TestPlan plan, VendorDatabase param, bool isEdit, bool isResubmit)
        {
            AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormVendorName, param.VendorName!, isEdit, true);
            AutomationHelper.CheckIsEditSelect(driver, VendorListXPath.FormVendorType, param.VendorType!, isEdit, true);

            AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormCity, param.City!, isEdit, true);
            AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormAddress, param.Address!, isEdit, true);

            AutomationHelper.CheckIsEditFillPhone(driver, VendorListXPath.FormPhone1_Front, VendorListXPath.FormPhone1_Main, VendorListXPath.FormPhone1_Back, param.Phone1!, isEdit, true);

            AutomationHelper.CheckIsEditFillPhone(driver, VendorListXPath.FormPhone2_Front, VendorListXPath.FormPhone2_Main, VendorListXPath.FormPhone2_Back, param.Phone2!, isEdit, false);

            AutomationHelper.CheckIsEditFillFax(driver, VendorListXPath.FormFaxNo_Front, VendorListXPath.FormFaxNo_Main, param.Fax!, isEdit, false);

            AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormEmail, param.Email!, isEdit, true);

            AutomationHelper.CheckIsEditSelect(driver, VendorListXPath.FormCountry, param.Country!, isEdit, true);

            AutomationHelper.ScrollToElement(driver, VendorListXPath.FormAttachmentBtnUpload);

            if (param.Country.Equals("indonesia", StringComparison.CurrentCultureIgnoreCase))
            {
                if (param.IsTaxableEntrepreneur)
                {
                    AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormTaxableTrue);
                    AutomationHelper.UploadFile(driver, VendorListXPath.FormBtnSPPKPAttachment, param.SPPKPPath);
                }
                else
                {
                    AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormTaxableFalse);
                }
            }
            else
            {
                if (param.IsHaveCertificateOfDomicile)
                {
                    AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormCoDTrue);
                    AutomationHelper.UploadFile(driver, VendorListXPath.FormBtnCoDAttachment, param.CoDAttachmentPath);
                    AutomationHelper.CheckIsEditFillFieldDate(driver, VendorListXPath.FormCoDExpDate, param.CoDExpiryDate!, isEdit, true);
                }
                else
                {
                    AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormCoDFalse);
                }
            }

            AutomationHelper.ScrollToElement(driver, VendorListXPath.FormAttachmentBtnUpload);

            if (param.IsUnderWithholdingTax)
            {
                AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormUWTTrue);
            }
            else
            {
                AutomationHelper.ClickInteractableElement(driver, VendorListXPath.FormUWTFalse);
            }

            AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormRemarks, param.Remarks!, isEdit, false);

            if (!string.IsNullOrEmpty(param.AttachmentPath))
            {
                AutomationHelper.UploadFile(driver, VendorListXPath.FormAttachmentBtnSelectFile, param.AttachmentPath);
                AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormAttachmentDescription, param.AttachmentDescription!, isEdit, true);
                AutomationHelper.ClickElement(driver, VendorListXPath.FormAttachmentBtnUpload);
            }

            AutomationHelper.ScrollToElement(driver, VendorListXPath.FormBtnNext);

            if (param.IsHaveVendorTaxId)
            {
                AutomationHelper.ClickInteractableElementUsingJS(driver, VendorListXPath.FormTaxIdTrue);

                AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormTaxIdNo, param.TaxIdNo!, isEdit, true);
                AutomationHelper.CheckIsEditFillField(driver, VendorListXPath.FormTaxIdName, param.TaxIdName!, isEdit, true);
                AutomationHelper.CheckIsEditFillField2(driver, VendorListXPath.FormTaxIdAddress, param.TaxIdAddress!, isEdit, true);
                AutomationHelper.UploadFile(driver, VendorListXPath.FormTaxIdBtnAttachment, param.TaxIdAttachmentPath);
            }
            else
            {
                AutomationHelper.ClickInteractableElementUsingJS(driver, VendorListXPath.FormTaxIdFalse);
                AutomationHelper.CheckIsEditFillField2(driver, VendorListXPath.FormTaxIdRemarks, param.TaxIdRemarks!, isEdit, true);
            }

            AutomationHelper.ClickElement(driver, VendorListXPath.FormBtnNext);
            string err = AutomationHelper.CheckAlertErrorVendorDatabase(driver);
            if (!string.IsNullOrEmpty(err))
            {
                throw new Exception(err);
            }
        }

        static void FillPhone(EdgeDriver driver, string XPathFront, string XPathMain, string XPathBack, string param)
        {
            var phones = param.Split('-');
            AutomationHelper.FillElement(driver, XPathFront, phones[0]);
            AutomationHelper.FillElement(driver, XPathMain, phones[1]);
            AutomationHelper.FillElement(driver, XPathBack, phones[2]);
        }

        private void SearchVendor<T>(EdgeDriver driver, AppConfig cfg, TestPlan plan, T param, bool isAfterSubmit = false, bool isTestCase = true) where T : class
        {
            VendorDatabaseSearch search = new();

            if (typeof(T) == typeof(VendorDatabaseSearch))
            {
                search = param as VendorDatabaseSearch ?? new();
            }
            else
            {
                VendorDatabase vendorDatabase = param as VendorDatabase ?? new();
                search = vendorDatabase.SearchVendor;
            }

            try
            {
                if (!isAfterSubmit)
                    OpenVendor(driver, cfg.Url);

                AutomationHelper.FillElementNonMandatory(driver, VendorListXPath.VendorName, search.VendorName);

                AutomationHelper.ClickElement(driver, VendorListXPath.BtnSearch);
                Thread.Sleep(2000);

                plan.Status = StatusConstant.Success;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
            }
            catch (Exception ex)
            {
                plan.Status = StatusConstant.Fail;
                plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : {ex.Message}";
                driver.Dispose();
            }
            finally
            {
                if (isTestCase)
                {
                    utils.WriteAutomationResult(cfg, plan);
                }
            }
        }

        static void OpenVendor(EdgeDriver driver, string baseUrl)
        {
            try
            {
                driver.Navigate().GoToUrl($"{baseUrl}/Vendor/VendorList");
            }
            catch (Exception ex)
            {
                throw new Exception($"Open Vendor Database is Fail : {ex.Message}");
            }
        }
    }
}
