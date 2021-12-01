using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Controls;

namespace WorkerToExel.Validations
{
    class FielldAndRusRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string field = (string)value;
            if (string.IsNullOrEmpty(field))
            {
                return new ValidationResult(false,
                    "Поле не может быть пустым");
            }
            else if (!Regex.IsMatch(field, @"^[A-Za-z0-9]+$"))
            {
                return new ValidationResult(false,
                    "Поле должно содержать только буквы латиницы и цифры");
            }
            return new ValidationResult(true, null);
        }
    }
}
