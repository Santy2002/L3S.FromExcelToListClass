using System;
using System.ComponentModel.DataAnnotations;


namespace FromExcelToListClass.CustomValidators
{
    public class BoolCustomValidator : ValidationAttribute
    {
        private string TrueValue { get; }
        private string FalseValue { get; }

        public BoolCustomValidator(string trueValue, string falseValue)
        {
            TrueValue = trueValue.ToLower();
            FalseValue = falseValue.ToLower();
        }

        public bool BoolCustomTryParse(string value, out bool boolValue)
        {
            if (value.Equals(TrueValue, StringComparison.InvariantCultureIgnoreCase))
            {
                boolValue = true;
                return true;
            }

            if (value.Equals(FalseValue, StringComparison.InvariantCultureIgnoreCase))
            {
                boolValue = false;
                return true;
            }

            boolValue = false;
            return false;
        }
    }
}
