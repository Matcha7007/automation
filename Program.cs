using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using System.Text;
using EPaymentVoucher.Models;
using EPaymentVoucher.Utilities;
using EPaymentVoucher.Actions.MyTaskList;
using EPaymentVoucher.Actions.PaymentVoucher;
using EPaymentVoucher.Actions.Vendor;
using EPaymentVoucher.Actions.BankConfirmation;

#region Configuration
var serviceProvider = new ServiceCollection()
	.AddSingleton<IMyTaskList, MyTaskList>()
	.AddSingleton<IHelperService, HelperService>()
	.AddSingleton<IUnitTestHandler, UnitTestHandler>()
	.AddSingleton<IVendor, Vendor>()
	.AddSingleton<IPVReimbursement, PVReimbursement>()
	.AddSingleton<IPaymentVoucher, PaymentVoucher>()
	.AddSingleton<IPVAdvance, PVAdvance>()
	.AddSingleton<IPVSettlement, PVSettlement>()
	.AddSingleton<IPVVendor, PVVendor>()
	.AddSingleton<IPVRDD, PVRDD>()
	.AddSingleton<IPVSalesForce, PVSalesForce>()
	.AddSingleton<IPVAgencyCommission, PVAgencyCommission>()
	.AddSingleton<IPVPolicyHolder, PVPolicyHolder>()
	.AddSingleton<IPVInternalBank, PVInternalBank>()
	.AddSingleton<IPVSpecialCustomer, PVSpecialCustomer>()
	.AddSingleton<IPVBancassuranceBMI, PVBancassuranceBMI>()
	.AddSingleton<IPVBancassuranceBTN, PVBancassuranceBTN>()
	.AddSingleton<IBankConfirmation, BankConfirmation>()
	.BuildServiceProvider();

var unitTest = serviceProvider.GetService<IUnitTestHandler>();

var basePath = "E:\\Project\\Selenium\\Zurich - Automation Test\\Zurich.AutomationTest\\EPaymentVoucher";

string jsonString = File.ReadAllText($"{basePath}\\appsettings.json", Encoding.Default);
AppConfig cfg = JsonConvert.DeserializeObject<AppConfig>(jsonString)!;
cfg.ExcelConfigPath = $"{basePath}{cfg.ExcelConfigPath}";
cfg.ScreenCapturePath = $"{basePath}{cfg.ScreenCapturePath}";
GlobalConfig.Config = cfg;
#endregion

Console.WriteLine($"Running Unit Test");
Console.WriteLine("------------------------------------------------------");

unitTest!.RunUnitTest(cfg);

Console.WriteLine($"End of Unit Test Sequence");
Console.WriteLine("------------------------------------------------------");
