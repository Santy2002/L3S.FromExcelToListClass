using L3S.FromExcelToListClass.Enums;

namespace L3S.FromExcelToListClass.Interfaces
{
    public interface IExcelEntity
    {
        int Row { get; set; }
        Errors Error { get; set; }
        public string TotalErrors { get; set; }
    }
}
