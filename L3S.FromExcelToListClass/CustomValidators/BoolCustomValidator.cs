using System;
using System.ComponentModel.DataAnnotations;


namespace L3S.FromExcelToListClass.CustomValidators
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class CustomValidBoolEntryAttribute : Attribute
    {
        private string TrueValue { get; }
        private string FalseValue { get; }

        public CustomValidBoolEntryAttribute(string trueValue, string falseValue)
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
