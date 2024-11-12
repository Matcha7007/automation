using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using System.Runtime.InteropServices;
using InputSimulatorEx;
using InputSimulatorEx.Native;
using EPaymentVoucher.Locators.General;
using SeleniumExtras.WaitHelpers;
using EPaymentVoucher.Models.Excel;
using EPaymentVoucher.Models;
using EPaymentVoucher.Locators;
using EPaymentVoucher.Locators.MyTaskList;

namespace EPaymentVoucher.Utilities
{
	public static class AutomationHelper
	{
		// Import necessary Windows API functions
		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool SetForegroundWindow(IntPtr hWnd);

		public static void OpenPage(EdgeDriver driver, string url)
		{
			try
			{
				driver.Navigate().GoToUrl(url);
				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				throw new Exception($"Open url {url} is Fail : {ex.Message}");
			}
		}

		public static void HandleAlertJS(EdgeDriver driver, bool isAccept)
		{
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(GlobalConfig.Config.WaitElementInSecond));
				wait.Until(ExpectedConditions.AlertIsPresent());

				IAlert alert = driver.SwitchTo().Alert();

				if (alert != null)
				{
					if (isAccept)
					{
						alert.Accept();
					}
					else
					{
						alert.Dismiss();
					}

					Thread.Sleep(1000);
				}
			}
			catch (WebDriverTimeoutException)
			{
				Console.WriteLine("No alert appeared within the timeout period.");
			}
			catch (NoAlertPresentException)
			{
				Console.WriteLine("No alert is present.");
			}
		}

		public static void Login(EdgeDriver driver, User user, AppConfig cfg)
		{
			try
			{
				driver.Navigate().GoToUrl(cfg.Url);

				FillField(driver, LoginXPath.Username, user.Username);
				FillField(driver, LoginXPath.Password, user.Password);

				ClickElement(driver, LoginXPath.BtnLogin);

				Console.WriteLine($"Login berhasil");
				Console.WriteLine($"Hello, {user.Username}!");
				Console.WriteLine("------------------------------------------------------");
				driver.Manage().Window.Maximize();
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				driver.Dispose();
			}
		}
		public static void FillField(EdgeDriver driver, string xpath, string param, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				if (string.IsNullOrEmpty(param))
				{
					throw new Exception($"The parameter is mandatory for the element {xpath}.");
				}
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
				driver.FindElement(By.XPath(xpath)).SendKeys(param);
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception($"The element {xpath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception($"The element {xpath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception($"The element {xpath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static string CheckAlertErrorVendorDatabase(EdgeDriver driver)
		{
			try
			{
				Thread.Sleep(1000);
				var alerts = driver.FindElements(By.XPath($"{MTVendorXPath.AlertError}/span"));
				string msg = string.Empty;
				foreach (var alert in alerts)
				{
					string err = alert.Text;
					msg = err.Replace("\r\n", " - ").Replace("\n", " - ").Replace("\r", " - ");
					break;
				}
				return msg;
			}
			catch { return string.Empty; }
		}

		public static void ClickButton(EdgeDriver driver, string xpath, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
				driver.FindElement(By.XPath(xpath)).Click();
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception($"The element {xpath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception($"The element {xpath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception($"The element {xpath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static void CheckElementExist(EdgeDriver driver, string xpath, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
			}
			catch (WebDriverTimeoutException ex)
			{
				Console.WriteLine($"The element {xpath} was not clickable within the timeout period: " + ex.Message);
				throw new Exception($"The element {xpath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				Console.WriteLine($"The element {xpath} could not be found: " + ex.Message);
				throw new Exception($"The element {xpath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				Console.WriteLine($"The element {xpath} is not interactable: " + ex.Message);
				throw new Exception($"The element {xpath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				Console.WriteLine($"An unexpected error occurred: " + ex.Message);
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static bool FindElements(EdgeDriver driver, string xpath, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{	
				CheckElementExist(driver, xpath, time);
				return true;
			}
			catch
			{
				return false;
			}
		}

		public static string CheckErrorReadExcel(string error)
		{
			return error.Contains("Input string was not in a correct format or must be in numeric format.") ? $"{error} Test Case Id or Sequence cannot be empty." : error;
		}

		public static bool ValidateAlert(string alert, string param)
		{
			return alert.ToLower().Contains(param.ToLower());
		}

		public static string GetAlertMessage(EdgeDriver driver, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.AlertIsPresent());
				IAlert alert = driver.SwitchTo().Alert();
				return alert.Text;
			}
			catch (NoAlertPresentException)
			{
				Console.WriteLine("No alert is present.");
				return "No alert is present.";
			}
		}

		public static void ToBeClickableByXPath(EdgeDriver driver, int timeOut, string xPath)
		{
			WebDriverWait wait = new(driver, TimeSpan.FromSeconds(timeOut));
			IWebElement clickableElement = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(xPath)));
			clickableElement.Click();
		}

		public static void UploadFileNonMandatory(EdgeDriver driver, string xPathBtn, string filePath, string title = "Open")
		{
			if (!string.IsNullOrEmpty(filePath))
				UploadFile(driver, xPathBtn, filePath);
		}

		public static void UploadFile(EdgeDriver driver, string xPathBtn, string filePath, string title = "Open")
		{
			try
			{
				if (string.IsNullOrEmpty(filePath))
				{
					throw new Exception($"File path parameter is mandatory for the element {xPathBtn}.");
				}
				Thread.Sleep(1000);
				var btnUpload = driver.FindElement(By.XPath(xPathBtn));

				OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver);
				action.Click(btnUpload).Build().Perform();

				Thread.Sleep(TimeSpan.FromSeconds(3));

				var dialogHWnd = FindWindow(null!, title); // Here goes the title of the dialog window
				var setFocus = SetForegroundWindow(dialogHWnd);
				if (setFocus)
				{
					var sim = new InputSimulator();
					sim.Keyboard
						.KeyPress(VirtualKeyCode.RETURN)
						.Sleep(150)
						.TextEntry(filePath)
						.Sleep(1000)
						.KeyPress(VirtualKeyCode.RETURN);
				}
				Thread.Sleep(1000);
			}
			catch (Exception ex)
			{
				throw new Exception($"Fail Upload File : {ex.Message}");
			}
		}

		public static void FillTaxId(EdgeDriver driver, string xpath, string param, string title = "Open")
		{
			try
			{
				if (!long.TryParse(param, out long parsedParam))
				{
					throw new ArgumentException("Tax Id must be in the format number");
				}

				char[] charArray = param.ToCharArray();

				if (charArray.Length != 16)
				{
					throw new ArgumentException("Tax Id must be 16 digit");
				}

				var taxBoxs = driver.FindElements(By.XPath(xpath));
				int i = 0;
				foreach (var taxBox in taxBoxs)
				{
					taxBox.SendKeys(charArray[i].ToString());
					i++;
					Thread.Sleep(100);
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Fail Fill Tax Id : {ex.Message}");
			}
		}

		public static void ScrollToElement(EdgeDriver driver, string xPathElement)
		{
			try
			{
				CheckElementExist(driver, xPathElement);
				var element = driver.FindElement(By.XPath(xPathElement));
				((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", element);
			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		public static void ScrollToElement(EdgeDriver driver, IWebElement element)
		{
			((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", element);
		}

		public static void ClickInteractableElement(EdgeDriver driver, string xPathElement)
		{
			try
			{
				CheckElementExist(driver, xPathElement);
				OpenQA.Selenium.Interactions.Actions action = new(driver);
				var element = driver.FindElement(By.XPath(xPathElement));
				action.Click(element).Build().Perform();
			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		public static void ClickInteractableElementUsingJS(EdgeDriver driver, string xPathElement)
		{
			try
			{
				CheckElementExist(driver, xPathElement);

				IWebElement element = driver.FindElement(By.XPath(xPathElement));
				IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
				jsExecutor.ExecuteScript($"arguments[0].click()", element);
			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		public static void InjectValueToHiddenElement(EdgeDriver driver, string xPathElement, string value, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
			IWebElement hiddenElement = wait.Until(d => d.FindElement(By.XPath(xPathElement)));

			IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
			jsExecutor.ExecuteScript("arguments[0].value = arguments[1]; arguments[0].setAttribute('value', arguments[1]); arguments[0].dispatchEvent(new Event('input'));", hiddenElement, value);
		}
		public static void DatePickerNonMandatory(EdgeDriver driver, string xPath, string param)
		{
			if (!string.IsNullOrEmpty(param))
				DatePicker(driver, xPath, param);
		}
		public static void DatePicker(EdgeDriver driver, string xPathElement, string date)
		{
			try
			{
				if (string.IsNullOrEmpty(date))
				{
					throw new Exception($"The parameter date is mandatory for the element {xPathElement}.");
				}

				ClickInteractableElement(driver, xPathElement);

				if (!DateTime.TryParse(date, out DateTime parsedDate))
				{
					throw new ArgumentException("Date must be in the format MM/DD/YYYY with ' at first. Example : '1/21/1991");
				}

				int day = parsedDate.Day;
				int month = parsedDate.Month;
				int year = parsedDate.Year;

				try
				{
					while (true)
					{
						Thread.Sleep(250);
						string datePickerMonth = driver.FindElement(By.XPath(DatePickerXPath.DatePickerMonth)).Text;
						if (!int.TryParse(driver.FindElement(By.XPath(DatePickerXPath.DatePickerYear)).Text, out int dtYear))
						{
							throw new InvalidOperationException("Unable to parse year from date picker");
						}

						int dtMonth = MonthHelper(datePickerMonth);

						if (month == dtMonth && year == dtYear)
						{
							break;
						}

						if (year == dtYear)
						{
							if (dtMonth < month)
							{
								driver.FindElement(By.XPath(DatePickerXPath.Next)).Click();
							}
							else if (dtMonth > month)
							{
								driver.FindElement(By.XPath(DatePickerXPath.Previous)).Click();
							}
						}
						else
						{
							if (dtYear < year)
							{
								driver.FindElement(By.XPath(DatePickerXPath.Next)).Click();
							}
							else if (dtYear > year)
							{
								driver.FindElement(By.XPath(DatePickerXPath.Previous)).Click();
							}
						}
					}
					string dayXPath = DatePickerXPath.DatePickerDay.Replace("replace_day", day.ToString());
					ClickInteractableElement(driver, dayXPath);
				}
				catch
				{
					string formattedDay = day.ToString("D2");
					string formattedMonth = month.ToString("D2");
					string formattedYear = year.ToString();
					string newDate = $"{formattedMonth}/{formattedDay}/{formattedYear}";
					InjectValueToHiddenElement(driver, xPathElement, newDate);
				}

			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		private static int MonthHelper(string month)
		{
			int intMonth = 0;
			switch (month)
			{
				case "January":
					intMonth = 1;
					break;
				case "February":
					intMonth = 2;
					break;
				case "March":
					intMonth = 3;
					break;
				case "April":
					intMonth = 4;
					break;
				case "May":
					intMonth = 5;
					break;
				case "June":
					intMonth = 6;
					break;
				case "July":
					intMonth = 7;
					break;
				case "August":
					intMonth = 8;
					break;
				case "September":
					intMonth = 9;
					break;
				case "October":
					intMonth = 10;
					break;
				case "November":
					intMonth = 11;
					break;
				case "December":
					intMonth = 12;
					break;
			}
			return intMonth;
		}

		public static void InjectValueToElementAndSetAttributeClass(EdgeDriver driver, string xPathElement, string value, string attribute)
		{
			IWebElement element = driver.FindElement(By.XPath(xPathElement));
			IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;

			jsExecutor.ExecuteScript("arguments[0].setAttribute('class', arguments[1]);", element, attribute);

			jsExecutor.ExecuteScript("arguments[0].value = arguments[1];", element, value);
		}

		public static void CheckIsEditFillField(EdgeDriver driver, string xPath, string param, bool isEdit, bool isMandatory)
		{
			try
			{
				CheckElementExist(driver, xPath);
				var ele = driver.FindElement(By.XPath(xPath));
				if (isEdit && !string.IsNullOrEmpty(param))
					ele.Clear();

				if (!string.IsNullOrEmpty(param) && (isMandatory || isEdit || !isMandatory))
					ele.SendKeys(param);

			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}
		public static void CheckIsEditFillField2(EdgeDriver driver, string xPath, string param, bool isEdit, bool isMandatory)
		{
			try
			{
				var ele = driver.FindElement(By.XPath(xPath));
				if (isEdit && !string.IsNullOrEmpty(param))
					ele.Clear();

				if (!string.IsNullOrEmpty(param) && (isMandatory || isEdit || !isMandatory))
					ele.SendKeys(param);

			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		public static void CheckIsEditFillFieldDate(EdgeDriver driver, string xPath, string param, bool isEdit, bool isMandatory)
		{
			try
			{
				var ele = driver.FindElement(By.XPath(xPath));
				if (isEdit && !string.IsNullOrEmpty(param))
					ele.Clear();

				if (!string.IsNullOrEmpty(param) && (isMandatory || isEdit || !isMandatory))
					DatePicker(driver, xPath, param);

			}
			catch (Exception ex) { throw new Exception(ex.Message); }
		}

		public static void CheckIsEditSelect(EdgeDriver driver, string xPath, string param, bool isEdit, bool isMandatory, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				if (string.IsNullOrEmpty(param) && isMandatory && !isEdit)
				{
					throw new Exception($"The parameter is mandatory for the element {xPath}.");
				}

				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xPath)));
				var type = driver.FindElement(By.XPath(xPath));
				Thread.Sleep(500);
				var typeOptions = type.FindElements(By.TagName("option"));
				Thread.Sleep(500);

				if (!string.IsNullOrEmpty(param) && (isMandatory || isEdit || !isMandatory))
				{
					foreach (var option in typeOptions)
					{
						if (option.Text.Trim().Contains(param.Trim(), StringComparison.OrdinalIgnoreCase))
						{
							option.Click();
						}
					}
				}
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception($"The element {xPath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception($"The element {xPath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception($"The element {xPath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static void SelectElementNonMandatory(EdgeDriver driver, string xPath, string param, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			if (!string.IsNullOrEmpty(param))
				SelectElement(driver, xPath, param, time);
		}
		public static void SelectElement(EdgeDriver driver, string xPath, string param, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				if (string.IsNullOrEmpty(param))
				{
					throw new Exception($"The parameter is mandatory for the element {xPath}.");
				}

				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xPath)));
				var type = driver.FindElement(By.XPath(xPath));
				Thread.Sleep(500);
				var typeOptions = type.FindElements(By.TagName("option"));
				Thread.Sleep(500);

				if (typeOptions.Count > 0)
				{
					bool optionFound = false;

					foreach (var option in typeOptions)
					{
						if (option.Text.Trim().Contains(param.Trim(), StringComparison.OrdinalIgnoreCase))
						{
							option.Click();
							optionFound = true;
							break;
						}
					}
					
					if (!optionFound)
					{
						throw new Exception($"The option '{param}' was not found in the element {xPath}.");
					}
				}
				else
				{
					throw new Exception($"No options were found in the element {xPath}.");
				}
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception($"The element {xPath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception($"The element {xPath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception($"The element {xPath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static void ValidateNumeric(string value)
		{
			try
			{
				int.Parse(value);
			}catch (Exception ex)
			{
				throw new Exception($"The value is not numeric. {ex.Message}");
			}
		}

		public static void FillElement(EdgeDriver driver, string xPath, string param, bool isNumeric = false, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				if (string.IsNullOrEmpty(param))
				{
					throw new Exception($"The parameter is mandatory for the element {xPath}.");
				}

				if (isNumeric)
				{
					ValidateNumeric(param);
				}

				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xPath)));
				IWebElement ele = driver.FindElement(By.XPath(xPath));
				ele.Clear();
				ele.SendKeys(param);

			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception($"The element {xPath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception($"The element {xPath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception($"The element {xPath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static void FillElementNonMandatory(EdgeDriver driver, string xPath, string param, bool isNumeric = false, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			if (!string.IsNullOrEmpty(param))
				FillElement(driver, xPath, param, isNumeric, time);
		}

		public static void ClickElement(EdgeDriver driver, string xPath, int time = -1)
		{
			time = time == -1 ? GlobalConfig.Config.WaitElementInSecond : time;
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(time));
				wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xPath)));
				ClickInteractableElement(driver, xPath);
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception($"The element {xPath} was not clickable within the timeout period: " + ex.Message);
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception($"The element {xPath} could not be found: " + ex.Message);
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception($"The element {xPath} is not interactable: " + ex.Message);
			}
			catch (Exception ex)
			{
				throw new Exception("An unexpected error occurred: " + ex.Message);
			}
		}

		public static void CheckIsEditFillPhone(EdgeDriver driver, string xPathFront, string xPathMain, string xPathBack, string param, bool isEdit, bool isMandatory)
		{
			var eleFront = driver.FindElement(By.XPath(xPathFront));
			var eleMain = driver.FindElement(By.XPath(xPathMain));
			var eleBack = driver.FindElement(By.XPath(xPathBack));

			if (isEdit && !string.IsNullOrEmpty(param))
			{
				eleFront.Clear();
				eleMain.Clear();
				eleBack.Clear();
			}

			if (!string.IsNullOrEmpty(param) && (isMandatory || isEdit || !isMandatory))
			{
				var strSplit = param.Split('-');
				try
				{
					int.Parse(strSplit[0]);
					int.Parse(strSplit[1]);
					int.Parse(strSplit[2]);
				}
				catch
				{
					throw new ArgumentException("Parameter is not numeric format. Example : 000-000000-000");
				}
				if (strSplit.Length > 2)
				{
					eleFront.SendKeys(strSplit[0]);
					eleMain.SendKeys(strSplit[1]);
					eleBack.SendKeys(strSplit[2]);
				}
				else
				{
					throw new ArgumentException("Parameter does not contain exactly three parts separated by hyphens. Example : 000-000000-000");
				}
			}

			Thread.Sleep(1000);
		}

		public static void CheckIsEditFillFax(EdgeDriver driver, string xPathFront, string xPathMain, string param, bool isEdit, bool isMandatory)
		{
			var eleFront = driver.FindElement(By.XPath(xPathFront));
			var eleMain = driver.FindElement(By.XPath(xPathMain));

			if (isEdit && !string.IsNullOrEmpty(param))
			{
				eleFront.Clear();
				eleMain.Clear();
			}

			if (!string.IsNullOrEmpty(param) && (isMandatory || isEdit || !isMandatory))
			{
				var strSplit = param.Split('-');
				try
				{
					int.Parse(strSplit[0]);
					int.Parse(strSplit[1]);
				}
				catch
				{
					throw new ArgumentException("Parameter is not numeric format.");
				}
				if (strSplit.Length > 1)
				{
					eleFront.SendKeys(strSplit[0]);
					eleMain.SendKeys(strSplit[1]);
				}
				else
				{
					throw new ArgumentException("Parameter does not contain exactly two parts separated by hyphens.");
				}
			}

			Thread.Sleep(1000);
		}
	}
}
