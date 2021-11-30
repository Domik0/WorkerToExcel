using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace WorkerToExel.Validations
{
    class EmailRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string email = (string)value;
            if (!Regex.IsMatch(email, @"^([a-z][a-z0-9]{4,19})((@gmail\.com)|(@mail\.ru)|(@bk\.ru)|(@yandex\.ru)|(@outlook\.com))$"))
            {
                return new ValidationResult(false,
                    "Некорректный email");
            }
            return new ValidationResult(true, null);
        }
    }
}
