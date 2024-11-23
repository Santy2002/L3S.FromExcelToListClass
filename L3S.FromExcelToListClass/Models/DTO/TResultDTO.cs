namespace L3S.FromExcelToListClass.Models.DTO
{
    public class TResultDTO<T> : ResultDTO where T : class
    {
        public List<T> ListResultT { get; set; } = new();
    }
}
