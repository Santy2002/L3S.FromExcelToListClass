using System.ComponentModel.DataAnnotations;

namespace L3S.FromExcelToListClass.CustomValidators
{
    public class RequiredButNullableAttribute : ValidationAttribute
    {
        public RequiredButNullableAttribute()
        {
            
        }

        public override bool IsValid(object value)
        {
            return base.IsValid(value);
        }
    }
}
