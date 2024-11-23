using NPOI.SS.UserModel;

namespace L3S.FromExcelToListClass.Models.DTO
{
    public class SheetResultDTO : ResultDTO
    {
        public ISheet Sheet { get; set; }
    }
}
