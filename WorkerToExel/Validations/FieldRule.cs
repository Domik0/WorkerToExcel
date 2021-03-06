using System.Globalization;
using System.Windows.Controls;

namespace WorkerToExel.Validations
{
    class FieldRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string field = (string)value;
            if (string.IsNullOrEmpty(field))
            {
                return new ValidationResult(false,
                    "Поле не может быть пустым");
            }
            return new ValidationResult(true, null);
        }
    }
}
