using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Controls;

namespace WorkerToExel.Validations
{
    class EmailRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string email = (string)value;
            if (!Regex.IsMatch(email,
                @"^([a-z][a-z0-9]{4,19})((@gmail\.com)|(@mail\.ru)|(@bk\.ru)|(@yandex\.ru)
                                            |(@outlook\.com)|(@akbars\.ru)|(@akbarsdigital\.ru))$"))
            {
                return new ValidationResult(false,
                    "Некорректный Email");
            }
            return new ValidationResult(true, null);
        }
    }
}
