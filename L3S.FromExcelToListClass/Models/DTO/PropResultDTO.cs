using NPOI.HSSF.Record.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace L3S.FromExcelToListClass.Models.DTO
{
    public class PropResultDTO<T> : ResultDTO where T : class
    {
        private T _valor;

        public PropResultDTO()
        {
        }

        public PropResultDTO(T valorInicial)
        {
            _valor = valorInicial;
        }
        public T Objeto
        {
            get { return _valor; }
            set { _valor = value; }
        }
    }
}
