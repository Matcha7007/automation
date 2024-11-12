namespace EPaymentVoucher.Locators.General
{
	public class DatePickerXPath
	{
		public const string Header = "//div[contains(@class, 'ui-datepicker-header')]";
		public const string Previous = $"{Header}//a[@title='Prev']";
		public const string Next = $"{Header}//a[@title='Next']";
		public const string DatePickerMonth = $"{Header}//span[@class='ui-datepicker-month']";
		public const string DatePickerYear = $"{Header}//span[@class='ui-datepicker-year']";
		public const string DatePickerDay = $"//td[@data-handler='selectDay']/a[@data-date='replace_day']";
	}
}
