using L3S.FromExcelToListClass.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace L3S.FromExcelToListClass.Interface
{
    public interface IExcelEntity
    {
        int Row { get; set; }
        Errors Error { get; set; }
        public string TotalErrors { get; set; }
    }
}
